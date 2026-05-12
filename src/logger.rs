use anyhow::Result;
use chrono::Local;
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::io::Write as IoWrite;
use std::path::Path;

#[derive(Debug, Clone, Serialize)]
pub struct ChangeRecord {
    pub timestamp: String,
    pub session_id: String,
    pub device_instance_id: String,
    pub device_display_name: String,
    pub reg_path: String,
    pub key_name: String,
    pub before_value: Option<Value>,
    pub after_value: Option<Value>,
    pub success: bool,
    pub error_message: Option<String>,
    pub backup_files: Vec<String>,
}

#[derive(Serialize)]
struct LogEntry<'a> {
    timestamp: &'a str,
    session_id: &'a str,
    device: DeviceInfo<'a>,
    change: ChangeInfo<'a>,
    backup_files: &'a [String],
}

#[derive(Serialize)]
struct DeviceInfo<'a> {
    instance_id: &'a str,
    display_name: &'a str,
}

#[derive(Serialize)]
struct ChangeInfo<'a> {
    reg_path: &'a str,
    key_name: &'a str,
    before_value: &'a Option<Value>,
    after_value: &'a Option<Value>,
    success: bool,
    error_message: &'a Option<String>,
}

// ──────────────────────────────────────────────
// セッション一覧（JSONL 読み込み）
// ──────────────────────────────────────────────

#[derive(Deserialize)]
struct StoredEntry {
    timestamp: String,
    session_id: String,
    device: StoredDevice,
    change: StoredChange,
    backup_files: Vec<String>,
}

#[derive(Deserialize)]
struct StoredDevice {
    instance_id: String,
    display_name: String,
}

#[derive(Deserialize)]
struct StoredChange {
    reg_path: String,
    key_name: String,
    before_value: Option<Value>,
    after_value: Option<Value>,
    success: bool,
    error_message: Option<String>,
}

pub struct SessionSummary {
    pub session_id: String,
    pub timestamp: String,
    pub records: Vec<ChangeRecord>,
}

pub fn list_sessions(log_dir: &Path) -> Vec<SessionSummary> {
    let rd = match std::fs::read_dir(log_dir) {
        Ok(d) => d,
        Err(_) => return Vec::new(),
    };

    let mut paths: Vec<_> = rd
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|e| e.to_str()) == Some("jsonl"))
        .collect();
    paths.sort();

    let mut order: Vec<String> = Vec::new();
    let mut groups: std::collections::HashMap<String, SessionSummary> = Default::default();

    for path in &paths {
        let Ok(content) = std::fs::read_to_string(path) else { continue };
        for line in content.lines().filter(|l| !l.trim().is_empty()) {
            let Ok(e) = serde_json::from_str::<StoredEntry>(line) else { continue };
            let sid = e.session_id.clone();
            if !groups.contains_key(&sid) {
                order.push(sid.clone());
                groups.insert(sid.clone(), SessionSummary {
                    session_id: sid.clone(),
                    timestamp: e.timestamp.clone(),
                    records: Vec::new(),
                });
            }
            groups.get_mut(&sid).unwrap().records.push(ChangeRecord {
                timestamp: e.timestamp,
                session_id: e.session_id,
                device_instance_id: e.device.instance_id,
                device_display_name: e.device.display_name,
                reg_path: e.change.reg_path,
                key_name: e.change.key_name,
                before_value: e.change.before_value,
                after_value: e.change.after_value,
                success: e.change.success,
                error_message: e.change.error_message,
                backup_files: e.backup_files,
            });
        }
    }

    // 新しいセッション順
    order.sort_by(|a, b| groups[b].timestamp.cmp(&groups[a].timestamp));
    order.into_iter().filter_map(|id| groups.remove(&id)).collect()
}

// ──────────────────────────────────────────────
// ログ書き込み
// ──────────────────────────────────────────────

pub fn log_changes(records: &[ChangeRecord], log_dir: &Path) -> Result<()> {
    if records.is_empty() {
        return Ok(());
    }
    std::fs::create_dir_all(log_dir)?;
    let date_str = Local::now().format("%Y%m%d").to_string();
    let now_str = Local::now().format("%Y-%m-%dT%H:%M:%S%.3f%:z").to_string();

    let jsonl_path = log_dir.join(format!("changes_{date_str}.jsonl"));
    let mut jsonl = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(&jsonl_path)?;

    let log_path = log_dir.join(format!("changes_{date_str}.log"));
    let mut logf = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(&log_path)?;

    for r in records {
        let entry = LogEntry {
            timestamp: &r.timestamp,
            session_id: &r.session_id,
            device: DeviceInfo {
                instance_id: &r.device_instance_id,
                display_name: &r.device_display_name,
            },
            change: ChangeInfo {
                reg_path: &r.reg_path,
                key_name: &r.key_name,
                before_value: &r.before_value,
                after_value: &r.after_value,
                success: r.success,
                error_message: &r.error_message,
            },
            backup_files: &r.backup_files,
        };
        writeln!(jsonl, "{}", serde_json::to_string(&entry)?)?;

        let status = if r.success { "OK" } else { "NG" };
        let err = r.error_message.as_deref().unwrap_or("");
        writeln!(
            logf,
            "{now_str} [{status}] {} | {}\\{} | {:?} -> {:?} | {err}",
            r.device_display_name, r.reg_path, r.key_name,
            r.before_value, r.after_value,
        )?;
    }
    Ok(())
}
