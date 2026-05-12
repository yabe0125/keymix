use anyhow::Result;
use serde_json::{json, Value};
use std::io::{self, Write as IoWrite};
use std::path::{Path, PathBuf};
use std::process::Command;
use winreg::enums::{HKEY_LOCAL_MACHINE, KEY_SET_VALUE};
use winreg::RegKey;

use crate::devices::{KeyboardDevice, Layout, PhysicalKeyboard, KEYBOARD_CLASS_GUID};
use crate::logger::ChangeRecord;

// ──────────────────────────────────────────────
// レイアウト定数
// ──────────────────────────────────────────────

pub const I8042PRT_PARAMS_PATH: &str =
    r"SYSTEM\CurrentControlSet\Services\i8042prt\Parameters";

pub struct LayoutSpec {
    pub keyboard_type: u32,
    pub keyboard_subtype: u32,
    pub layer_driver_jpn: &'static str,
    pub name: &'static str,
}

pub const LAYOUT_US: LayoutSpec = LayoutSpec {
    keyboard_type: 4,
    keyboard_subtype: 0,
    layer_driver_jpn: "kbdus.dll",
    name: "US配列 (101/102キー)",
};

pub const LAYOUT_JIS: LayoutSpec = LayoutSpec {
    keyboard_type: 7,
    keyboard_subtype: 2,
    layer_driver_jpn: "kbd106.dll",
    name: "JIS配列 (106/109キー)",
};

pub fn layout_spec(layout: Layout) -> &'static LayoutSpec {
    match layout {
        Layout::Us => &LAYOUT_US,
        Layout::Jis => &LAYOUT_JIS,
    }
}

// ──────────────────────────────────────────────
// レジストリユーティリティ
// ──────────────────────────────────────────────

pub fn read_dword(path: &str, name: &str) -> Option<u32> {
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);
    hklm.open_subkey(path).ok()?.get_value(name).ok()
}

pub fn read_string(path: &str, name: &str) -> Option<String> {
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);
    hklm.open_subkey(path).ok()?.get_value(name).ok()
}

fn open_or_create_for_write(path: &str) -> io::Result<RegKey> {
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);
    match hklm.open_subkey_with_flags(path, KEY_SET_VALUE) {
        Ok(k) => Ok(k),
        Err(e) if e.kind() == io::ErrorKind::NotFound => {
            hklm.create_subkey(path).map(|(k, _)| k)
        }
        Err(e) => Err(e),
    }
}

// ──────────────────────────────────────────────
// i8042prt システム設定
// ──────────────────────────────────────────────

pub fn get_i8042prt_override() -> Option<Layout> {
    Layout::from_dword(read_dword(I8042PRT_PARAMS_PATH, "OverrideKeyboardType")?)
}

pub fn i8042prt_status_label(override_val: Option<Layout>) -> &'static str {
    match override_val {
        None => "接続済みキーボード レイアウトを使用する（自動）",
        Some(Layout::Us) => "英語キーボード (101/102キー) に固定",
        Some(Layout::Jis) => "日本語キーボード (106/109キー) に固定",
    }
}

fn backup_i8042prt(timestamp: &str, backup_dir: &Path) -> Result<PathBuf> {
    std::fs::create_dir_all(backup_dir)?;
    let out_file = backup_dir.join(format!("backup_{timestamp}_i8042prt.reg"));
    let full_path = format!("HKLM\\{I8042PRT_PARAMS_PATH}");
    let output = Command::new("reg")
        .args(["export", &full_path, out_file.to_str().unwrap_or(""), "/y"])
        .output()?;
    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        anyhow::bail!("reg export 失敗: {}", stderr.trim());
    }
    Ok(out_file)
}

pub fn apply_system_keyboard_layout(
    layout_key: Option<Layout>,
    timestamp: &str,
    session_id: &str,
    backup_dir: &Path,
    no_backup: bool,
) -> Vec<ChangeRecord> {
    let full_path = format!("HKEY_LOCAL_MACHINE\\{I8042PRT_PARAMS_PATH}");
    println!("  変更対象レジストリキー:");
    println!("    {full_path}");
    match layout_key {
        None => {
            println!("    OverrideKeyboardType      → (削除)");
            println!("    OverrideKeyboardSubtype   → (削除)");
            println!("    LayerDriver JPN           → (削除)");
        }
        Some(layout) => {
            let spec = layout_spec(layout);
            println!("    OverrideKeyboardType    = {}", spec.keyboard_type);
            println!("    OverrideKeyboardSubtype = {}", spec.keyboard_subtype);
            println!("    LayerDriver JPN         = {}", spec.layer_driver_jpn);
        }
    }
    print!("  続行しますか？ [y/N]: ");
    io::stdout().flush().ok();
    let mut input = String::new();
    if io::stdin().read_line(&mut input).is_err() || !input.trim().eq_ignore_ascii_case("y") {
        println!("  キャンセルしました。");
        return Vec::new();
    }

    let mut backup_files: Vec<String> = Vec::new();

    if !no_backup {
        print!("  バックアップを作成しています...");
        io::stdout().flush().ok();
        match backup_i8042prt(timestamp, backup_dir) {
            Ok(f) => {
                backup_files.push(f.to_string_lossy().into_owned());
                println!(" [OK]");
                println!("    {}", f.display());
            }
            Err(e) => {
                println!(" [失敗]");
                println!("  エラー: {e}");
                println!("  安全のため、レジストリの変更を中断します。");
                return Vec::new();
            }
        }
    }

    println!("  レジストリを変更しています...");
    let mut records = Vec::new();

    match layout_key {
        None => {
            for val_name in ["OverrideKeyboardType", "OverrideKeyboardSubtype", "LayerDriver JPN"] {
                let before = read_dword(I8042PRT_PARAMS_PATH, val_name)
                    .map(|v| json!(v))
                    .or_else(|| read_string(I8042PRT_PARAMS_PATH, val_name).map(|s| json!(s)));

                let (success, err_msg) = if before.is_none() {
                    print!("    {val_name}: [削除済み（スキップ）]");
                    (true, None)
                } else {
                    match open_or_create_for_write(I8042PRT_PARAMS_PATH) {
                        Ok(key) => match key.delete_value(val_name) {
                            Ok(()) => { print!("    {val_name}: 削除"); (true, None) }
                            Err(e) if e.kind() == io::ErrorKind::NotFound => {
                                print!("    {val_name}: [削除済み（スキップ）]");
                                (true, None)
                            }
                            Err(e) => { print!("    {val_name}:"); (false, Some(e.to_string())) }
                        },
                        Err(e) => { print!("    {val_name}:"); (false, Some(e.to_string())) }
                    }
                };

                if success { println!(" [OK]"); } else if let Some(ref m) = err_msg { println!(" [失敗: {m}]"); }

                records.push(ChangeRecord {
                    timestamp: timestamp.to_string(),
                    session_id: session_id.to_string(),
                    device_instance_id: "i8042prt".to_string(),
                    device_display_name: "システム設定 (i8042prt)".to_string(),
                    reg_path: I8042PRT_PARAMS_PATH.to_string(),
                    key_name: val_name.to_string(),
                    before_value: before,
                    after_value: None,
                    success,
                    error_message: err_msg,
                    backup_files: backup_files.clone(),
                });
            }
        }
        Some(layout) => {
            let spec = layout_spec(layout);
            let dword_ops: &[(&str, u32)] = &[
                ("OverrideKeyboardType", spec.keyboard_type),
                ("OverrideKeyboardSubtype", spec.keyboard_subtype),
            ];
            for &(val_name, after_val) in dword_ops {
                let before = read_dword(I8042PRT_PARAMS_PATH, val_name).map(|v| json!(v));
                let (success, err_msg) = match open_or_create_for_write(I8042PRT_PARAMS_PATH) {
                    Ok(key) => match key.set_value(val_name, &after_val) {
                        Ok(()) => (true, None),
                        Err(e) => (false, Some(e.to_string())),
                    },
                    Err(e) => (false, Some(e.to_string())),
                };
                println!("    {val_name} = {after_val} {}", if success { "[OK]" } else { "[失敗]" });
                records.push(ChangeRecord {
                    timestamp: timestamp.to_string(),
                    session_id: session_id.to_string(),
                    device_instance_id: "i8042prt".to_string(),
                    device_display_name: "システム設定 (i8042prt)".to_string(),
                    reg_path: I8042PRT_PARAMS_PATH.to_string(),
                    key_name: val_name.to_string(),
                    before_value: before,
                    after_value: Some(json!(after_val)),
                    success,
                    error_message: err_msg,
                    backup_files: backup_files.clone(),
                });
            }
            let before = read_string(I8042PRT_PARAMS_PATH, "LayerDriver JPN").map(|s| json!(s));
            let dll = spec.layer_driver_jpn;
            let (success, err_msg) = match open_or_create_for_write(I8042PRT_PARAMS_PATH) {
                Ok(key) => match key.set_value("LayerDriver JPN", &dll) {
                    Ok(()) => (true, None),
                    Err(e) => (false, Some(e.to_string())),
                },
                Err(e) => (false, Some(e.to_string())),
            };
            println!("    LayerDriver JPN = {dll} {}", if success { "[OK]" } else { "[失敗]" });
            records.push(ChangeRecord {
                timestamp: timestamp.to_string(),
                session_id: session_id.to_string(),
                device_instance_id: "i8042prt".to_string(),
                device_display_name: "システム設定 (i8042prt)".to_string(),
                reg_path: I8042PRT_PARAMS_PATH.to_string(),
                key_name: "LayerDriver JPN".to_string(),
                before_value: before,
                after_value: Some(json!(dll)),
                success,
                error_message: err_msg,
                backup_files: backup_files.clone(),
            });
        }
    }

    records
}

// ──────────────────────────────────────────────
// バックアップ（デバイス単位）
// ──────────────────────────────────────────────

pub fn create_registry_backup(
    physical_kb: &PhysicalKeyboard,
    timestamp: &str,
    backup_dir: &Path,
) -> Result<Vec<PathBuf>> {
    std::fs::create_dir_all(backup_dir)?;

    let container_short = physical_kb
        .container_id
        .trim_matches(|c| c == '{' || c == '}')
        .replace('-', "");
    let container_short = &container_short[..container_short.len().min(8)];

    let mut backup_files = Vec::new();
    let mut errors = Vec::new();
    let mut seen_hw_keys = std::collections::HashSet::new();

    for col in &physical_kb.collections {
        if col.bus != "HID" || !seen_hw_keys.insert(col.hw_key.clone()) {
            continue;
        }
        let reg_key = format!("HKLM\\SYSTEM\\CurrentControlSet\\Enum\\{}\\{}", col.bus, col.hw_key);
        let safe_name: String = col.hw_key.chars()
            .map(|c| if "\\/:*?\"<>|{}".contains(c) { '_' } else { c })
            .take(30)
            .collect();
        let out_file = backup_dir.join(format!(
            "backup_{timestamp}_{container_short}_enum_{safe_name}.reg"
        ));
        let out = Command::new("reg")
            .args(["export", &reg_key, out_file.to_str().unwrap_or(""), "/y"])
            .output()?;
        if out.status.success() {
            backup_files.push(out_file);
        } else {
            errors.push(format!("reg export 失敗 ({reg_key})"));
        }
    }

    // Class ドライバのバックアップ
    let class_key = format!(
        "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Class\\{{{KEYBOARD_CLASS_GUID}}}"
    );
    let class_file = backup_dir.join(format!(
        "backup_{timestamp}_{container_short}_class.reg"
    ));
    let out = Command::new("reg")
        .args(["export", &class_key, class_file.to_str().unwrap_or(""), "/y"])
        .output()?;
    if out.status.success() {
        backup_files.push(class_file);
    } else {
        errors.push("reg export 失敗 (Class)".to_string());
    }

    if !errors.is_empty() {
        anyhow::bail!("{}", errors.join("\n"));
    }
    Ok(backup_files)
}

// ──────────────────────────────────────────────
// デバイス単位のレジストリ書き込み
// ──────────────────────────────────────────────

fn write_device_parameters(
    col: &KeyboardDevice,
    spec: &LayoutSpec,
    timestamp: &str,
    session_id: &str,
    backup_files: &[String],
) -> Vec<ChangeRecord> {
    let dp_path = format!("{}\\Device Parameters", col.enum_path());
    let mut records = Vec::new();

    for (val_name, after_val) in [
        ("OverrideKeyboardType", spec.keyboard_type),
        ("OverrideKeyboardSubtype", spec.keyboard_subtype),
    ] {
        let before = read_dword(&dp_path, val_name).map(|v| json!(v));
        let (success, err_msg) = match open_or_create_for_write(&dp_path) {
            Ok(key) => match key.set_value(val_name, &after_val) {
                Ok(()) => (true, None),
                Err(e) => (false, Some(e.to_string())),
            },
            Err(e) => (false, Some(e.to_string())),
        };
        records.push(ChangeRecord {
            timestamp: timestamp.to_string(),
            session_id: session_id.to_string(),
            device_instance_id: col.instance_id(),
            device_display_name: col.display_name.clone(),
            reg_path: dp_path.clone(),
            key_name: val_name.to_string(),
            before_value: before,
            after_value: Some(json!(after_val)),
            success,
            error_message: err_msg,
            backup_files: backup_files.to_vec(),
        });
    }
    records
}

fn write_class_driver(
    col: &KeyboardDevice,
    spec: &LayoutSpec,
    timestamp: &str,
    session_id: &str,
    backup_files: &[String],
) -> Vec<ChangeRecord> {
    if col.driver_class_instance.is_empty() {
        return Vec::new();
    }
    let inst_num = match col.driver_class_instance.rsplit('\\').next() {
        Some(s) => s,
        None => return Vec::new(),
    };
    let class_path = format!(
        "SYSTEM\\CurrentControlSet\\Control\\Class\\{{{KEYBOARD_CLASS_GUID}}}\\{inst_num}"
    );
    let mut records = Vec::new();

    for (val_name, after_val) in [
        ("KeyboardType", spec.keyboard_type),
        ("KeyboardSubType", spec.keyboard_subtype),
    ] {
        let before = read_dword(&class_path, val_name).map(|v| json!(v));
        let (success, err_msg) = match open_or_create_for_write(&class_path) {
            Ok(key) => match key.set_value(val_name, &after_val) {
                Ok(()) => (true, None),
                Err(e) => (false, Some(e.to_string())),
            },
            Err(e) => (false, Some(format!("クラスドライバ: {e}"))),
        };
        records.push(ChangeRecord {
            timestamp: timestamp.to_string(),
            session_id: session_id.to_string(),
            device_instance_id: col.instance_id(),
            device_display_name: col.display_name.clone(),
            reg_path: class_path.clone(),
            key_name: val_name.to_string(),
            before_value: before,
            after_value: Some(json!(after_val)),
            success,
            error_message: err_msg,
            backup_files: backup_files.to_vec(),
        });
    }
    records
}

pub fn apply_layout_to_device(
    physical_kb: &PhysicalKeyboard,
    spec: &LayoutSpec,
    timestamp: &str,
    session_id: &str,
    backup_dir: &Path,
    no_backup: bool,
) -> Vec<ChangeRecord> {
    let mut backup_file_paths: Vec<String> = Vec::new();

    if !no_backup {
        print!("  バックアップを作成しています...");
        io::stdout().flush().ok();
        match create_registry_backup(physical_kb, timestamp, backup_dir) {
            Ok(files) => {
                backup_file_paths = files.iter().map(|p| p.to_string_lossy().into_owned()).collect();
                println!(" [OK]");
                for f in &files {
                    println!("    {}", f.display());
                }
            }
            Err(e) => {
                println!(" [失敗]");
                println!("  エラー: {e}");
                println!("  安全のため、レジストリの変更を中断します。");
                return Vec::new();
            }
        }
    }

    println!("  レジストリを変更しています...");
    let mut all_records = Vec::new();

    for col in &physical_kb.collections {
        if col.bus != "HID" {
            continue;
        }

        let hw_short = col.hw_key.chars().take(50).collect::<String>();
        print!("    {hw_short}\\...\\Device Parameters");
        io::stdout().flush().ok();

        let dp_recs = write_device_parameters(col, spec, timestamp, session_id, &backup_file_paths);
        if dp_recs.iter().all(|r| r.success) {
            println!(" [OK]");
        } else {
            let msg = dp_recs.iter().find(|r| !r.success)
                .and_then(|r| r.error_message.as_deref())
                .unwrap_or("エラー");
            println!(" [一部失敗: {msg}]");
        }
        all_records.extend(dp_recs);

        if !col.driver_class_instance.is_empty() {
            let inst_num = col.driver_class_instance.rsplit('\\').next().unwrap_or("?");
            print!("    Class\\{{...}}\\{inst_num}");
            io::stdout().flush().ok();

            let cls_recs = write_class_driver(col, spec, timestamp, session_id, &backup_file_paths);
            if cls_recs.is_empty() {
                println!(" [スキップ]");
            } else if cls_recs.iter().all(|r| r.success) {
                println!(" [OK]");
            } else {
                let msg = cls_recs.iter().find(|r| !r.success)
                    .and_then(|r| r.error_message.as_deref())
                    .unwrap_or("エラー");
                println!(" [スキップ: {msg}]");
            }
            all_records.extend(cls_recs);
        }
    }

    all_records
}

// ──────────────────────────────────────────────
// セッション復元
// ──────────────────────────────────────────────

pub fn restore_session(
    records: &[ChangeRecord],
    new_timestamp: &str,
    new_session_id: &str,
) -> Vec<ChangeRecord> {
    let mut results = Vec::new();

    for rec in records.iter().filter(|r| r.success) {
        let current_val = read_dword(&rec.reg_path, &rec.key_name)
            .map(|v| json!(v))
            .or_else(|| read_string(&rec.reg_path, &rec.key_name).map(|s| json!(s)));

        let label_target = match &rec.before_value {
            None => "(削除)".to_string(),
            Some(v) => v.to_string(),
        };
        let label_current = current_val
            .as_ref()
            .map(|v| v.to_string())
            .unwrap_or_else(|| "(なし)".to_string());
        print!("    {}: {} → {label_target}", rec.key_name, label_current);
        io::stdout().flush().ok();

        let (success, err_msg) = match &rec.before_value {
            None => {
                match open_or_create_for_write(&rec.reg_path) {
                    Ok(key) => match key.delete_value(&rec.key_name) {
                        Ok(()) => (true, None),
                        Err(e) if e.kind() == io::ErrorKind::NotFound => (true, None),
                        Err(e) => (false, Some(e.to_string())),
                    },
                    Err(e) => (false, Some(e.to_string())),
                }
            }
            Some(Value::Number(n)) => {
                let v = n.as_u64().unwrap_or(0) as u32;
                match open_or_create_for_write(&rec.reg_path) {
                    Ok(key) => match key.set_value(&rec.key_name, &v) {
                        Ok(()) => (true, None),
                        Err(e) => (false, Some(e.to_string())),
                    },
                    Err(e) => (false, Some(e.to_string())),
                }
            }
            Some(Value::String(s)) => {
                match open_or_create_for_write(&rec.reg_path) {
                    Ok(key) => match key.set_value(&rec.key_name, s) {
                        Ok(()) => (true, None),
                        Err(e) => (false, Some(e.to_string())),
                    },
                    Err(e) => (false, Some(e.to_string())),
                }
            }
            Some(_) => (false, Some("対応していない値型です".to_string())),
        };

        if success {
            println!(" [OK]");
        } else {
            println!(" [失敗: {}]", err_msg.as_deref().unwrap_or(""));
        }

        results.push(ChangeRecord {
            timestamp: new_timestamp.to_string(),
            session_id: new_session_id.to_string(),
            device_instance_id: rec.device_instance_id.clone(),
            device_display_name: rec.device_display_name.clone(),
            reg_path: rec.reg_path.clone(),
            key_name: rec.key_name.clone(),
            before_value: current_val,
            after_value: rec.before_value.clone(),
            success,
            error_message: err_msg,
            backup_files: Vec::new(),
        });
    }

    results
}
