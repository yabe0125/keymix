use embed_manifest::{embed_manifest, new_manifest};
use embed_manifest::manifest::ExecutionLevel;

fn main() {
    if std::env::var("CARGO_CFG_TARGET_OS").as_deref() == Ok("windows") {
        embed_manifest(
            new_manifest("keyboard-config")
                .requested_execution_level(ExecutionLevel::RequireAdministrator),
        )
        .expect("マニフェストの埋め込みに失敗しました");
    }
    println!("cargo:rerun-if-changed=build.rs");
}
