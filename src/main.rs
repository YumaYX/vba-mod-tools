use clap::{Parser, Subcommand};
use std::fs::{self, File};
use std::io::{self, BufRead, BufReader, Write};
use std::path::Path;

#[derive(Parser)]
#[command(name = "bas-tools", version = "1.0", about = "VBA file splitter/merger with CRLF/LF conversion")]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// 分割: CRLF -> LF に変換して vba/ へ
    Split { 
        input: String, 
        #[arg(short, long, default_value = "vba")] out_dir: String 
    },
    /// 結合: LF -> CRLF に変換して module.bas へ (vba/ は削除)
    Concat { 
        #[arg(short, long, default_value = "vba")] src_dir: String, 
        #[arg(short, long, default_value = "module.bas")] output: String 
    },
}

fn main() -> io::Result<()> {
    let cli = Cli::parse();

    match cli.command {
        Commands::Split { input, out_dir } => {
            if Path::new(&out_dir).exists() {
                fs::remove_dir_all(&out_dir)?;
            }
            fs::create_dir_all(&out_dir)?;

            let file = File::open(&input)?;
            let reader = BufReader::new(file);

            let mut current_file: Option<File> = None;
            let mut buffer = Vec::new();

            for line in reader.lines() {
                let line = line?;
                if line.starts_with("'#########") {
                    // 分割時は LF で保存
                    save_buffer(&mut current_file, &mut buffer, "\n")?;
                    
                    let func_name = line.split_whitespace().nth(1).unwrap_or("unknown");
                    let path = Path::new(&out_dir).join(format!("{}.bas", func_name));
                    current_file = Some(File::create(path)?);
                }
                if current_file.is_some() {
                    buffer.push(line);
                }
            }
            save_buffer(&mut current_file, &mut buffer, "\n")?;
            fs::remove_file(input)?;
            println!("Split完了: LFに変換して {} に保存しました。", out_dir);
        }
        Commands::Concat { src_dir, output } => {
            // --- フォルダ存在チェックを追加 ---
            let src_path = Path::new(&src_dir);
            if !src_path.exists() {
                eprintln!("エラー: フォルダ '{}' が見つかりません。先に split を実行してください。", src_dir);
                std::process::exit(1);
            }

            let mut out = File::create(&output)?;
            let mut entries: Vec<_> = fs::read_dir(src_path)?
                .filter_map(|e| e.ok())
                .collect();
            
            if entries.is_empty() {
                eprintln!("エラー: フォルダ '{}' は空です。", src_dir);
                std::process::exit(1);
            }

            // アルファベット・数字順にソート
            entries.sort_by_key(|e| e.path());

            for entry in entries {
                let content = fs::read_to_string(entry.path())?;
                let trimmed = content.trim_end();
                if !trimmed.is_empty() {
                    // 結合時は CRLF (\r\n) で出力
                    write!(out, "{}", trimmed.replace("\n", "\r\n"))?;
                    write!(out, "\r\n\r\n")?; 
                }
            }

            // 結合が終わったらフォルダを削除
            //fs::remove_dir_all(src_path)?;
            println!("成功: CRLF形式で '{}' を作成しました", output);
        }
    }
    Ok(())
}

fn save_buffer(file: &mut Option<File>, buffer: &mut Vec<String>, newline: &str) -> io::Result<()> {
    if let Some(mut f) = file.take() {
        while buffer.last().map_or(false, |s| s.trim().is_empty()) {
            buffer.pop();
        }
        for line in buffer.iter() {
            // 指定した改行コードで書き出し
            write!(f, "{}{}", line, newline)?;
        }
        buffer.clear();
    }
    Ok(())
}
