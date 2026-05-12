use regex::Regex;
use std::collections::HashMap;
use std::sync::OnceLock;
use winreg::enums::HKEY_LOCAL_MACHINE;
use winreg::RegKey;

pub const KEYBOARD_CLASS_GUID: &str = "4D36E96B-E325-11CE-BFC1-08002BE10318";
const INTERNAL_CONTAINER_ID: &str = "{00000000-0000-0000-ffff-ffffffffffff}";
const BT_HID_UUIDS: &[&str] = &["00001812", "00001124"];
const SCAN_BUSES: &[&str] = &["HID", "ACPI", "BTHENUM", "BTH", "USB"];
const KEYBOARD_SERVICES: &[&str] = &["kbdhid", "i8042prt"];

static VENDOR_NAMES: OnceLock<HashMap<&'static str, &'static str>> = OnceLock::new();

fn vendor_names() -> &'static HashMap<&'static str, &'static str> {
    VENDOR_NAMES.get_or_init(|| {
        HashMap::from([
            ("3434", "Keychron"),
            ("068E", "HyperX"),
            ("046D", "Logitech"),
            ("045E", "Microsoft"),
            ("1532", "Razer"),
            ("1B1C", "Corsair"),
            ("04D9", "Holtek"),
            ("04F2", "Chicony"),
            ("04CA", "Lite-On"),
            ("0461", "Primax"),
            ("05AC", "Apple"),
            ("04E8", "Samsung"),
            ("03F0", "HP"),
            ("413C", "Dell"),
            ("17EF", "Lenovo"),
            ("0B05", "ASUS ROG"),
            ("1038", "SteelSeries"),
        ])
    })
}

static USB_RE: OnceLock<Regex> = OnceLock::new();
static BT_RE: OnceLock<Regex> = OnceLock::new();

fn usb_re() -> &'static Regex {
    USB_RE.get_or_init(|| {
        Regex::new(r"(?i)VID_([0-9A-Fa-f]{4}).*?PID_([0-9A-Fa-f]{4})").unwrap()
    })
}

fn bt_re() -> &'static Regex {
    BT_RE.get_or_init(|| {
        Regex::new(r"(?i)VID&[0-9A-Fa-f]{2}([0-9A-Fa-f]{4}).*?PID&([0-9A-Fa-f]{4})").unwrap()
    })
}

// ──────────────────────────────────────────────
// 公開型
// ──────────────────────────────────────────────

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Layout {
    Us,
    Jis,
}

impl Layout {
    pub fn label(self) -> &'static str {
        match self {
            Layout::Us => "US配列",
            Layout::Jis => "JIS配列",
        }
    }

    pub fn from_dword(kt: u32) -> Option<Layout> {
        match kt {
            4 => Some(Layout::Us),
            7 => Some(Layout::Jis),
            _ => None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ConnectionType {
    Internal,
    Bluetooth,
    Usb,
    Unknown,
}

impl ConnectionType {
    pub fn label(&self) -> &'static str {
        match self {
            ConnectionType::Internal => "内蔵",
            ConnectionType::Bluetooth => "Bluetooth",
            ConnectionType::Usb => "USB",
            ConnectionType::Unknown => "不明",
        }
    }
}

#[derive(Debug, Clone)]
pub struct KeyboardDevice {
    pub bus: String,
    pub hw_key: String,
    pub instance: String,
    pub container_id: String,
    pub display_name: String,
    pub connection_type: ConnectionType,
    pub driver_class_instance: String,
    pub current_layout: Option<Layout>,
}

impl KeyboardDevice {
    pub fn instance_id(&self) -> String {
        format!("{}\\{}\\{}", self.bus, self.hw_key, self.instance)
    }

    pub fn enum_path(&self) -> String {
        format!(
            "SYSTEM\\CurrentControlSet\\Enum\\{}\\{}\\{}",
            self.bus, self.hw_key, self.instance
        )
    }
}

#[derive(Debug, Clone)]
pub struct PhysicalKeyboard {
    pub container_id: String,
    pub display_name: String,
    pub connection_type: ConnectionType,
    pub collections: Vec<KeyboardDevice>,
}

impl PhysicalKeyboard {
    pub fn current_layout(&self) -> Option<Layout> {
        self.collections.iter().find_map(|c| c.current_layout)
    }
}

// ──────────────────────────────────────────────
// デバイス列挙
// ──────────────────────────────────────────────

pub fn enumerate_keyboard_devices() -> Vec<KeyboardDevice> {
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);
    let mut devices = Vec::new();

    for &bus in SCAN_BUSES {
        let bus_path = format!("SYSTEM\\CurrentControlSet\\Enum\\{bus}");
        let bus_key = match hklm.open_subkey(&bus_path) {
            Ok(k) => k,
            Err(_) => continue,
        };

        let hw_keys: Vec<String> = bus_key.enum_keys().filter_map(|r| r.ok()).collect();

        for hw_key_name in hw_keys {
            let hw_path = format!("{bus_path}\\{hw_key_name}");
            let instances = {
                match hklm.open_subkey(&hw_path) {
                    Ok(k) => k.enum_keys().filter_map(|r| r.ok()).collect::<Vec<_>>(),
                    Err(_) => continue,
                }
            };

            for inst_name in instances {
                let inst_path = format!("{hw_path}\\{inst_name}");
                let inst_key = match hklm.open_subkey(&inst_path) {
                    Ok(k) => k,
                    Err(_) => continue,
                };

                let service: String = inst_key.get_value("Service").unwrap_or_default();
                let service_lower = service.to_ascii_lowercase();
                if !KEYBOARD_SERVICES.contains(&service_lower.as_str()) {
                    continue;
                }

                let container_id: String = inst_key.get_value("ContainerID").unwrap_or_default();
                let driver: String = inst_key.get_value("Driver").unwrap_or_default();
                let device_desc: String = inst_key.get_value("DeviceDesc").unwrap_or_default();

                let conn_type = classify_device(bus, &hw_key_name, &container_id, &service_lower);
                let display_name = build_display_name(&hw_key_name, &device_desc);
                let enum_path = format!(
                    "SYSTEM\\CurrentControlSet\\Enum\\{bus}\\{hw_key_name}\\{inst_name}"
                );
                let current_layout = read_current_layout(&enum_path, &driver);

                devices.push(KeyboardDevice {
                    bus: bus.to_string(),
                    hw_key: hw_key_name.clone(),
                    instance: inst_name,
                    container_id,
                    display_name,
                    connection_type: conn_type,
                    driver_class_instance: driver,
                    current_layout,
                });
            }
        }
    }

    devices
}

pub fn group_by_physical_device(devices: Vec<KeyboardDevice>) -> Vec<PhysicalKeyboard> {
    let mut groups: HashMap<String, PhysicalKeyboard> = HashMap::new();

    for dev in devices {
        let key = if !dev.container_id.is_empty() {
            dev.container_id.to_ascii_lowercase()
        } else {
            dev.hw_key.clone()
        };

        let entry = groups.entry(key).or_insert_with(|| PhysicalKeyboard {
            container_id: dev.container_id.clone(),
            display_name: dev.display_name.clone(),
            connection_type: dev.connection_type.clone(),
            collections: Vec::new(),
        });
        entry.collections.push(dev);
    }

    let mut result: Vec<PhysicalKeyboard> = groups.into_values().collect();
    // 外付け → 内蔵の順
    result.sort_by_key(|kb| if kb.connection_type == ConnectionType::Internal { 1 } else { 0 });
    result
}

// ──────────────────────────────────────────────
// デバイス分類
// ──────────────────────────────────────────────

fn classify_device(
    bus: &str,
    hw_key: &str,
    container_id: &str,
    service: &str,
) -> ConnectionType {
    if container_id.eq_ignore_ascii_case(INTERNAL_CONTAINER_ID) {
        return ConnectionType::Internal;
    }
    if bus == "ACPI" {
        return ConnectionType::Internal;
    }
    let hw_lower = hw_key.to_ascii_lowercase();
    if hw_lower.starts_with("converteddevice") {
        return ConnectionType::Internal;
    }
    if BT_HID_UUIDS.iter().any(|uuid| hw_lower.contains(uuid)) {
        return ConnectionType::Bluetooth;
    }
    if matches!(bus, "BTHENUM" | "BTH" | "BTHLE") {
        return ConnectionType::Bluetooth;
    }
    if matches!(service, "bthhidenum" | "bthhid" | "bthledeviceenum") {
        return ConnectionType::Bluetooth;
    }
    if usb_re().is_match(hw_key) {
        return ConnectionType::Usb;
    }
    if bus == "USB" {
        return ConnectionType::Usb;
    }
    ConnectionType::Unknown
}

fn extract_vid_pid(hw_key: &str) -> Option<(String, String)> {
    if let Some(cap) = usb_re().captures(hw_key) {
        return Some((cap[1].to_ascii_uppercase(), cap[2].to_ascii_uppercase()));
    }
    // BT VID format has a 2-char source prefix before the actual vendor ID
    if let Some(cap) = bt_re().captures(hw_key) {
        return Some((cap[1].to_ascii_uppercase(), cap[2].to_ascii_uppercase()));
    }
    None
}

fn build_display_name(hw_key: &str, device_desc: &str) -> String {
    if let Some((vid, pid)) = extract_vid_pid(hw_key) {
        if let Some(&vendor) = vendor_names().get(vid.as_str()) {
            return format!("{vendor} (PID:{pid})");
        }
        return format!("Keyboard VID:{vid} PID:{pid}");
    }
    let desc = if device_desc.contains(';') {
        device_desc.rsplit(';').next().unwrap_or(device_desc).trim()
    } else {
        device_desc.trim()
    };
    if !desc.is_empty() {
        return desc.to_string();
    }
    hw_key.chars().take(40).collect()
}

// ──────────────────────────────────────────────
// 現在のレイアウト読み取り
// ──────────────────────────────────────────────

pub fn read_current_layout(enum_path: &str, driver: &str) -> Option<Layout> {
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);

    let dp_path = format!("{enum_path}\\Device Parameters");
    if let Ok(key) = hklm.open_subkey(&dp_path) {
        if let Ok(kt) = key.get_value::<u32, _>("OverrideKeyboardType") {
            return Layout::from_dword(kt);
        }
    }

    if !driver.is_empty() {
        let inst_num = driver.rsplit('\\').next()?;
        let class_path = format!(
            "SYSTEM\\CurrentControlSet\\Control\\Class\\{{{KEYBOARD_CLASS_GUID}}}\\{inst_num}"
        );
        if let Ok(key) = hklm.open_subkey(&class_path) {
            if let Ok(kt) = key.get_value::<u32, _>("KeyboardType") {
                return Layout::from_dword(kt);
            }
        }
    }

    None
}
