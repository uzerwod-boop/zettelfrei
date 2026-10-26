// Zettelfrei als Programm für Windows und macOS.
// Die Oberfläche ist dieselbe wie im Browser (web/dist). Rust übernimmt nur,
// was ein Browser-Fenster nicht kann: Dateien über den „Speichern unter“-Dialog sichern.
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use base64::{engine::general_purpose::STANDARD, Engine};
use tauri_plugin_dialog::DialogExt;

/// Fragt nach dem Speicherort und schreibt die Datei. Gibt den Pfad zurück
/// oder None, wenn der Dialog abgebrochen wurde.
#[tauri::command]
async fn save_file(app: tauri::AppHandle, name: String, data: String) -> Result<Option<String>, String> {
    let bytes = STANDARD.decode(data).map_err(|e| e.to_string())?;
    let Some(path) = app.dialog().file().set_file_name(&name).blocking_save_file() else {
        return Ok(None);
    };
    let path = path.into_path().map_err(|e| e.to_string())?;
    std::fs::write(&path, bytes).map_err(|e| e.to_string())?;
    Ok(Some(path.display().to_string()))
}

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![save_file])
        .run(tauri::generate_context!())
        .expect("Zettelfrei konnte nicht gestartet werden");
}
