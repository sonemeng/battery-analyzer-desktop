use std::fs;
use std::path::Path;
use std::process::Command;
use serde::{Deserialize, Serialize};

#[derive(Debug, Serialize, Deserialize)]
struct FileInfo {
    name: String,
    path: String,
    size: u64,
    is_excel: bool,
    last_modified: String,
}

#[derive(Debug, Serialize, Deserialize)]
struct ProcessConfig {
    // 基础配置
    input_folder: String,
    output_folder: String,

    // 异常检测配置
    outlier_method: String,
    boxplot_threshold_discharge: f64,
    boxplot_threshold_efficiency: f64,
    zscore_threshold_discharge: f64,
    zscore_threshold_efficiency: f64,
    zscore_mad_constant: f64,

    // 其他配置（可选，使用默认值）
    #[serde(default)]
    reference_channel_method: String,
    #[serde(default)]
    verbose: bool,
    #[serde(default)]
    enable_progress_bar: bool,
}

// Tauri命令：读取目录文件
#[tauri::command]
fn read_directory(path: String) -> Result<Vec<FileInfo>, String> {
    let dir_path = Path::new(&path);
    if !dir_path.exists() {
        return Err("文件夹不存在".to_string());
    }

    let mut files = Vec::new();

    match fs::read_dir(dir_path) {
        Ok(entries) => {
            for entry in entries {
                if let Ok(entry) = entry {
                    let file_name = entry.file_name().to_string_lossy().to_string();
                    let file_path = entry.path().to_string_lossy().to_string();
                    let is_excel = file_name.ends_with(".xlsx") || file_name.ends_with(".xls");

                    // 只包含Excel文件
                    if is_excel {
                        let metadata = entry.metadata().map_err(|e| format!("读取文件元数据失败: {}", e))?;
                        let size = metadata.len();
                        let last_modified = format!("{:?}", metadata.modified().unwrap_or(std::time::SystemTime::UNIX_EPOCH));

                        files.push(FileInfo {
                            name: file_name,
                            path: file_path,
                            size,
                            is_excel,
                            last_modified,
                        });
                    }
                }
            }
        }
        Err(e) => return Err(format!("读取文件夹失败: {}", e)),
    }

    Ok(files)
}

// Tauri命令：处理电池数据（调用Python模块）
#[tauri::command]
fn process_battery_data(config: ProcessConfig) -> Result<String, String> {
    // 检查输入文件夹是否存在
    if !Path::new(&config.input_folder).exists() {
        return Err("输入文件夹不存在".to_string());
    }

    // 创建输出文件夹（如果不存在）
    let output_folder = if config.output_folder.is_empty() {
        config.input_folder.clone()
    } else {
        config.output_folder.clone()
    };

    if let Err(e) = fs::create_dir_all(&output_folder) {
        return Err(format!("创建输出文件夹失败: {}", e));
    }

    // 构建Python脚本路径（相对于Tauri应用）
    let python_script = "../main.py";

    // 检查Python脚本是否存在
    if !Path::new(python_script).exists() {
        return Err(format!("Python脚本不存在: {}", python_script));
    }

    // 构建命令参数（传递所有配置）
    let mut cmd = Command::new("python");
    cmd.arg(python_script)
       .arg("--input_folder").arg(&config.input_folder)
       .arg("--output_folder").arg(&output_folder)
       .arg("--outlier_method").arg(&config.outlier_method)
       .arg("--reference_channel_method").arg(&config.reference_channel_method)
       .arg("--boxplot_threshold_discharge").arg(config.boxplot_threshold_discharge.to_string())
       .arg("--boxplot_threshold_efficiency").arg(config.boxplot_threshold_efficiency.to_string())
       .arg("--zscore_threshold_discharge").arg(config.zscore_threshold_discharge.to_string())
       .arg("--zscore_threshold_efficiency").arg(config.zscore_threshold_efficiency.to_string())
       .arg("--zscore_mad_constant").arg(config.zscore_mad_constant.to_string());

    // 添加可选参数
    if config.verbose {
        cmd.arg("--verbose");
    }

    // 执行Python脚本
    match cmd.output() {
        Ok(output) => {
            if output.status.success() {
                let stdout = String::from_utf8_lossy(&output.stdout);
                Ok(format!("✅ 数据处理完成！\n\n{}", stdout))
            } else {
                let stderr = String::from_utf8_lossy(&output.stderr);
                Err(format!("❌ Python脚本执行失败:\n{}", stderr))
            }
        }
        Err(e) => Err(format!("❌ 启动Python脚本失败: {}", e))
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
  tauri::Builder::default()
    .plugin(tauri_plugin_dialog::init())
    .plugin(tauri_plugin_fs::init())
    .invoke_handler(tauri::generate_handler![
        read_directory,
        process_battery_data
    ])
    .setup(|app| {
      if cfg!(debug_assertions) {
        app.handle().plugin(
          tauri_plugin_log::Builder::default()
            .level(log::LevelFilter::Info)
            .build(),
        )?;
      }
      Ok(())
    })
    .run(tauri::generate_context!())
    .expect("error while running tauri application");
}
