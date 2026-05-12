mod devices;
mod logger;
mod ops;

use chrono::Local;
use clap::{Parser, Subcommand};
use std::io::{self, Write as IoWrite};
use std::path::{Path, PathBuf};

use devices::{enumerate_keyboard_devices, group_by_physical_device, Layout, PhysicalKeyboard};
use logger::{list_sessions, log_changes, SessionSummary};
use ops::{apply_layout_to_device, apply_system_keyboard_layout, get_i8042prt_override,
          i8042prt_status_label, layout_spec, restore_session, LAYOUT_JIS, LAYOUT_US};

// ──────────────────────────────────────────────
// Windows API
// ──────────────────────────────────────────────

#[cfg(windows)]
fn is_admin() -> bool {
    extern "system" {
        fn IsUserAnAdmin() -> i32;
    }
    unsafe { IsUserAnAdmin() != 0 }
}

#[cfg(not(windows))]
fn is_admin() -> bool {
    false
}

#[cfg(windows)]
fn setup_console() {
    extern "system" {
        fn SetConsoleOutputCP(wCodePageID: u32) -> i32;
    }
    unsafe { SetConsoleOutputCP(65001) };
}

#[cfg(not(windows))]
fn setup_console() {}

// ──────────────────────────────────────────────
// CLI 定義
// ──────────────────────────────────────────────

#[derive(Parser)]
#[command(
    name = "keyboard-config",
    version = "1.1.0",
    about = "Windows キーボードレイアウト設定ツール"
)]
struct Args {
    #[command(subcommand)]
    command: Option<Cmd>,

    /// キーボード一覧を表示して終了
    #[arg(long)]
    list: bool,

    /// バックアップ保存先ディレクトリ
    #[arg(long, value_name = "DIR")]
    backup_dir: Option<PathBuf>,

    /// ログ保存先ディレクトリ
    #[arg(long, value_name = "DIR")]
    log_dir: Option<PathBuf>,

    /// バックアップをスキップ
    #[arg(long)]
    no_backup: bool,
}

#[derive(Subcommand)]
enum Cmd {
    /// 指定のデバイスのレイアウトを変更
    Set {
        /// デバイス番号（--list で確認）
        #[arg(value_name = "番号")]
        number: usize,
        /// レイアウト us または jis
        #[arg(value_name = "us|jis")]
        layout: String,
    },
    /// システムキーボード設定を変更（i8042prt）
    SetSystem {
        /// auto (接続済みを使用), us, jis
        #[arg(value_name = "auto|us|jis")]
        layout: String,
    },
    /// 直近の変更を元に戻す
    Restore {
        /// 復元するセッションID（省略時は対話選択）
        #[arg(value_name = "SESSION_ID")]
        session_id: Option<String>,
        /// セッション一覧を表示して終了
        #[arg(long)]
        list: bool,
    },
}

// ──────────────────────────────────────────────
// ユーティリティ
// ──────────────────────────────────────────────

fn default_backup_dir() -> PathBuf {
    dirs_sys_backup()
}

fn default_log_dir() -> PathBuf {
    dirs_sys_log()
}

#[cfg(windows)]
fn dirs_sys_backup() -> PathBuf {
    let local_app = std::env::var("LOCALAPPDATA").unwrap_or_else(|_| ".".to_string());
    PathBuf::from(local_app).join("keyboard-config").join("backups")
}

#[cfg(not(windows))]
fn dirs_sys_backup() -> PathBuf {
    PathBuf::from("backups")
}

#[cfg(windows)]
fn dirs_sys_log() -> PathBuf {
    let local_app = std::env::var("LOCALAPPDATA").unwrap_or_else(|_| ".".to_string());
    PathBuf::from(local_app).join("keyboard-config").join("logs")
}

#[cfg(not(windows))]
fn dirs_sys_log() -> PathBuf {
    PathBuf::from("logs")
}

fn make_session_id() -> String {
    format!("{:08x}", Local::now().timestamp() as u32)
}

fn make_timestamp() -> String {
    Local::now().format("%Y%m%d_%H%M%S").to_string()
}

fn prompt(msg: &str) -> String {
    print!("{msg}");
    io::stdout().flush().ok();
    let mut buf = String::new();
    io::stdin().read_line(&mut buf).ok();
    buf.trim().to_string()
}

fn layout_current_label(layout: Option<Layout>) -> &'static str {
    layout.map_or("未設定", Layout::label)
}

fn parse_layout_name(s: &str) -> Option<Layout> {
    match s.to_ascii_lowercase().as_str() {
        "us" => Some(Layout::Us),
        "jis" => Some(Layout::Jis),
        _ => None,
    }
}

// ──────────────────────────────────────────────
// 一覧表示
// ──────────────────────────────────────────────

fn print_keyboard_list(all_keyboards: &[PhysicalKeyboard], sys_override: Option<Layout>) {
    println!("\nキーボード一覧:");
    println!("  S. システム設定 (i8042prt): {}", i8042prt_status_label(sys_override));

    for (i, kb) in all_keyboards.iter().enumerate() {
        let conn = kb.connection_type.label();
        let layout = layout_current_label(kb.current_layout());
        println!("  {}. {}  [{}]  現在の設定: {}", i + 1, kb.display_name, conn, layout);
    }

    if all_keyboards.is_empty() {
        println!("  （キーボードが検出されませんでした）");
    }
}

// ──────────────────────────────────────────────
// セッション一覧・復元
// ──────────────────────────────────────────────

fn print_session_list(sessions: &[SessionSummary]) {
    if sessions.is_empty() {
        println!("  復元可能なセッションはありません。");
        return;
    }
    println!("\n復元可能なセッション:");
    for (i, s) in sessions.iter().enumerate() {
        let mut device_names: Vec<&str> = Vec::new();
        for r in s.records.iter().filter(|r| r.success) {
            if !device_names.contains(&r.device_display_name.as_str()) {
                device_names.push(&r.device_display_name);
            }
        }
        let count = s.records.iter().filter(|r| r.success).count();
        let devs = if device_names.is_empty() {
            "（変更なし）".to_string()
        } else {
            device_names.join(", ")
        };
        println!("  {:2}. {} [{}]  {}  ({count}件)", i + 1, s.timestamp, s.session_id, devs);
    }
}

fn do_restore(session: &SessionSummary, log_dir: &Path) {
    let restorables: Vec<_> = session.records.iter().filter(|r| r.success).collect();
    if restorables.is_empty() {
        println!("  このセッションに復元可能な変更はありません。");
        return;
    }

    println!("  復元する変更:");
    for r in &restorables {
        let val_target = match &r.before_value {
            None => "(削除)".to_string(),
            Some(v) => v.to_string(),
        };
        let val_now = r
            .after_value
            .as_ref()
            .map(|v| v.to_string())
            .unwrap_or_else(|| "(なし)".to_string());
        println!("    {} | {}: {} → {val_target}", r.device_display_name, r.key_name, val_now);
    }

    let confirm = prompt("\n続行しますか？ [y/N]: ");
    if confirm.to_ascii_lowercase() != "y" {
        println!("キャンセルしました。");
        return;
    }

    let timestamp = make_timestamp();
    let new_session_id = make_session_id();
    println!("\n  レジストリを復元しています...");
    let records = restore_session(&session.records, &timestamp, &new_session_id);

    if records.is_empty() {
        return;
    }

    let ok_count = records.iter().filter(|r| r.success).count();
    let ng_count = records.len() - ok_count;
    if ng_count == 0 {
        println!("\n完了。変更を有効にするにはデバイスの再接続または再起動が必要です。");
    } else {
        println!("\n一部の復元に失敗しました（成功: {ok_count} / 失敗: {ng_count}）。");
    }

    if let Err(e) = log_changes(&records, log_dir) {
        eprintln!("ログ書き込みエラー: {e}");
    }
}

fn run_restore_command(session_id_arg: Option<&str>, log_dir: &Path) -> i32 {
    let sessions = list_sessions(log_dir);

    let idx = if let Some(sid) = session_id_arg {
        match sessions.iter().position(|s| s.session_id == sid) {
            Some(i) => i,
            None => {
                eprintln!("エラー: セッション「{sid}」が見つかりません。");
                return 1;
            }
        }
    } else {
        print_session_list(&sessions);
        if sessions.is_empty() {
            return 1;
        }
        let input = prompt("\n> 復元するセッションの番号を入力 (q でキャンセル): ");
        if input.to_ascii_lowercase() == "q" {
            return 0;
        }
        match input.parse::<usize>() {
            Ok(n) if n >= 1 && n <= sessions.len() => n - 1,
            _ => {
                println!("無効な入力です。");
                return 1;
            }
        }
    };

    let session = &sessions[idx];
    println!("\nセッション [{}] {} を元に戻します", session.session_id, session.timestamp);
    do_restore(session, log_dir);
    0
}

// ──────────────────────────────────────────────
// 対話モード：システム設定サブメニュー
// ──────────────────────────────────────────────

fn run_system_setting(
    timestamp: &str,
    session_id: &str,
    backup_dir: &Path,
    log_dir: &Path,
    no_backup: bool,
) {
    println!("\nシステムキーボード設定 (i8042prt):");
    let current = get_i8042prt_override();
    println!("  現在: {}", i8042prt_status_label(current));
    println!();
    println!("  1. 接続済みキーボードレイアウトを使用する（自動）");
    println!("  2. 英語キーボード (101/102キー) に固定");
    println!("  3. 日本語キーボード (106/109キー) に固定");

    let input = prompt("\n> 番号を入力 (b で戻る): ");
    let layout_key: Option<Option<Layout>> = match input.as_str() {
        "1" => Some(None),
        "2" => Some(Some(Layout::Us)),
        "3" => Some(Some(Layout::Jis)),
        "b" | "B" => return,
        _ => {
            println!("無効な入力です。");
            return;
        }
    };

    let layout_key = layout_key.unwrap();
    let label = i8042prt_status_label(layout_key);
    println!("\n「{label}」に変更します。");

    let records = apply_system_keyboard_layout(
        layout_key,
        timestamp,
        session_id,
        backup_dir,
        no_backup,
    );

    if records.is_empty() {
        return;
    }

    let ok_count = records.iter().filter(|r| r.success).count();
    let ng_count = records.len() - ok_count;
    if ng_count == 0 {
        println!("\n完了。変更を有効にするには再起動が必要です。");
    } else {
        println!("\n一部の変更に失敗しました（成功: {ok_count} / 失敗: {ng_count}）。");
    }

    if let Err(e) = log_changes(&records, log_dir) {
        eprintln!("ログ書き込みエラー: {e}");
    }
}

// ──────────────────────────────────────────────
// 対話モード：メインループ
// ──────────────────────────────────────────────

fn run_interactive(
    all_keyboards: &[PhysicalKeyboard],
    backup_dir: &Path,
    log_dir: &Path,
    no_backup: bool,
) {
    let timestamp = make_timestamp();
    let session_id = make_session_id();

    loop {
        let sys_override = get_i8042prt_override();
        print_keyboard_list(all_keyboards, sys_override);

        let max_num = all_keyboards.len();
        let input = prompt("\n> 変更するキーボードの番号を入力 (S: システム設定, R: 元に戻す, q で終了): ");

        match input.to_ascii_lowercase().as_str() {
            "q" => break,
            "r" => {
                if !is_admin() {
                    println!("レジストリの変更には管理者権限が必要です。");
                    continue;
                }
                let sessions = list_sessions(log_dir);
                print_session_list(&sessions);
                if sessions.is_empty() {
                    continue;
                }
                let input_r = prompt("\n> 復元するセッションの番号を入力 (b で戻る): ");
                if input_r.to_ascii_lowercase() == "b" {
                    continue;
                }
                let idx = match input_r.parse::<usize>() {
                    Ok(n) if n >= 1 && n <= sessions.len() => n - 1,
                    _ => {
                        println!("無効な入力です。");
                        continue;
                    }
                };
                let session = &sessions[idx];
                println!("\nセッション [{}] {} を元に戻します", session.session_id, session.timestamp);
                do_restore(session, log_dir);
            }
            "s" => {
                if !is_admin() {
                    println!("システム設定の変更には管理者権限が必要です。");
                    continue;
                }
                run_system_setting(&timestamp, &session_id, backup_dir, log_dir, no_backup);
            }
            s => {
                let num: usize = match s.parse() {
                    Ok(n) if n >= 1 && n <= max_num => n,
                    _ => {
                        println!("無効な入力です。{max_num} の番号、または q を入力してください。");
                        continue;
                    }
                };

                if !is_admin() {
                    println!("レジストリの変更には管理者権限が必要です。");
                    continue;
                }

                let kb = &all_keyboards[num - 1];
                println!("\n選択: {} [{}]", kb.display_name, kb.connection_type.label());
                println!("\nレイアウトを選択:");
                println!("  1. US配列 (101/102キー)");
                println!("  2. JIS配列 (106/109キー)");

                let layout_input = prompt("\n> 番号を入力 (b で戻る): ");
                let spec = match layout_input.as_str() {
                    "1" => &LAYOUT_US,
                    "2" => &LAYOUT_JIS,
                    "b" | "B" => continue,
                    _ => {
                        println!("無効な入力です。");
                        continue;
                    }
                };

                println!("\n{}を{}に変更します。", kb.display_name, spec.name);

                let records = apply_layout_to_device(
                    kb,
                    spec,
                    &timestamp,
                    &session_id,
                    backup_dir,
                    no_backup,
                );

                if records.is_empty() {
                    continue;
                }

                let ok_count = records.iter().filter(|r| r.success).count();
                let ng_count = records.len() - ok_count;
                if ng_count == 0 {
                    println!("\n完了。変更を有効にするにはデバイスの再接続または再起動が必要です。");
                } else {
                    println!("\n一部の変更に失敗しました。");
                }

                if let Err(e) = log_changes(&records, log_dir) {
                    eprintln!("ログ書き込みエラー: {e}");
                }
            }
        }

        println!();
    }
}

// ──────────────────────────────────────────────
// 非対話コマンド
// ──────────────────────────────────────────────

fn run_set_command(
    all_keyboards: &[PhysicalKeyboard],
    number: usize,
    layout_name: &str,
    backup_dir: &Path,
    log_dir: &Path,
    no_backup: bool,
) -> i32 {
    if number < 1 || number > all_keyboards.len() {
        eprintln!(
            "エラー: デバイス番号 {number} は無効です（有効範囲: 1～{}）。",
            all_keyboards.len()
        );
        return 1;
    }

    let layout = match parse_layout_name(layout_name) {
        Some(l) => l,
        None => {
            eprintln!("エラー: レイアウト「{layout_name}」は無効です（us または jis を指定）。");
            return 1;
        }
    };

    let kb = &all_keyboards[number - 1];
    let spec = layout_spec(layout);
    let timestamp = make_timestamp();
    let session_id = make_session_id();

    println!("{}を{}に変更します。", kb.display_name, spec.name);

    let records = apply_layout_to_device(kb, spec, &timestamp, &session_id, backup_dir, no_backup);

    if records.is_empty() {
        return 1;
    }

    let ng_count = records.iter().filter(|r| !r.success).count();
    if ng_count == 0 {
        println!("完了。変更を有効にするにはデバイスの再接続または再起動が必要です。");
    } else {
        println!("一部の変更に失敗しました。");
    }

    if let Err(e) = log_changes(&records, log_dir) {
        eprintln!("ログ書き込みエラー: {e}");
    }

    if ng_count > 0 { 1 } else { 0 }
}

fn run_set_system_command(
    layout_name: &str,
    backup_dir: &Path,
    log_dir: &Path,
    no_backup: bool,
) -> i32 {
    let layout_key: Option<Layout> = match layout_name.to_ascii_lowercase().as_str() {
        "auto" => None,
        "us" => Some(Layout::Us),
        "jis" => Some(Layout::Jis),
        _ => {
            eprintln!("エラー: 「{layout_name}」は無効です（auto, us, jis のいずれかを指定）。");
            return 1;
        }
    };

    let timestamp = make_timestamp();
    let session_id = make_session_id();

    println!("システム設定を「{}」に変更します。", i8042prt_status_label(layout_key));

    let records =
        apply_system_keyboard_layout(layout_key, &timestamp, &session_id, backup_dir, no_backup);

    if records.is_empty() {
        return 1;
    }

    let ng_count = records.iter().filter(|r| !r.success).count();
    if ng_count == 0 {
        println!("完了。変更を有効にするには再起動が必要です。");
    } else {
        println!("一部の変更に失敗しました。");
    }

    if let Err(e) = log_changes(&records, log_dir) {
        eprintln!("ログ書き込みエラー: {e}");
    }

    if ng_count > 0 { 1 } else { 0 }
}

// ──────────────────────────────────────────────
// エントリポイント
// ──────────────────────────────────────────────

fn main() {
    setup_console();

    let args = Args::parse();

    let backup_dir = args.backup_dir.unwrap_or_else(default_backup_dir);
    let log_dir = args.log_dir.unwrap_or_else(default_log_dir);
    let no_backup = args.no_backup;

    println!("Keyboard Layout Configuration Tool");
    println!("====================================");

    if args.list {
        let devices = enumerate_keyboard_devices();
        let keyboards = group_by_physical_device(devices);
        let sys_override = get_i8042prt_override();
        print_keyboard_list(&keyboards, sys_override);
        return;
    }

    let admin_ok = is_admin();

    match args.command {
        Some(Cmd::Set { number, layout }) => {
            if !admin_ok {
                eprintln!("エラー: レジストリの変更には管理者権限が必要です。");
                std::process::exit(2);
            }
            let devices = enumerate_keyboard_devices();
            let keyboards = group_by_physical_device(devices);
            let code = run_set_command(&keyboards, number, &layout, &backup_dir, &log_dir, no_backup);
            std::process::exit(code);
        }
        Some(Cmd::SetSystem { layout }) => {
            if !admin_ok {
                eprintln!("エラー: システム設定の変更には管理者権限が必要です。");
                std::process::exit(2);
            }
            let code = run_set_system_command(&layout, &backup_dir, &log_dir, no_backup);
            std::process::exit(code);
        }
        Some(Cmd::Restore { session_id, list }) => {
            if list {
                let sessions = list_sessions(&log_dir);
                print_session_list(&sessions);
                return;
            }
            if !admin_ok {
                eprintln!("エラー: レジストリの変更には管理者権限が必要です。");
                std::process::exit(2);
            }
            let code = run_restore_command(session_id.as_deref(), &log_dir);
            std::process::exit(code);
        }
        None => {
            let devices = enumerate_keyboard_devices();
            let keyboards = group_by_physical_device(devices);
            run_interactive(&keyboards, &backup_dir, &log_dir, no_backup);
        }
    }
}
