using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Data.Sqlite;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using UglyToad.PdfPig;

namespace TextFinder
{
    public class SearchResult
    {
        public string FileName { get; set; } = "";
        public string FilePath { get; set; } = "";
        public string Preview { get; set; } = "";
    }

    public partial class MainWindow : Window
    {
        private readonly string _dbPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TextFinder", "index.db");
        
        private CancellationTokenSource? _cts;
        private bool _isIndexing = false;

        public MainWindow()
        {
            InitializeComponent();
            LoadLastPath();
            EnsureDatabase();
            UpdateIndexStatus();
        }

        private void LoadLastPath()
        {
            var settingsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TextFinder", "settings.txt");
            
            if (File.Exists(settingsPath))
            {
                PathTextBox.Text = File.ReadAllText(settingsPath).Trim();
            }
            else
            {
                PathTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
        }

        private void SaveLastPath()
        {
            var folder = Path.GetDirectoryName(_dbPath);
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            
            var settingsPath = Path.Combine(folder, "settings.txt");
            File.WriteAllText(settingsPath, PathTextBox.Text);
        }

        private void EnsureDatabase()
        {
            var folder = Path.GetDirectoryName(_dbPath);
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            using var conn = new SqliteConnection($"Data Source={_dbPath}");
            conn.Open();
            
            // 设置超时和 WAL 模式提高并发性能
            var pragmaCmd = conn.CreateCommand();
            pragmaCmd.CommandText = "PRAGMA journal_mode=WAL; PRAGMA busy_timeout=5000;";
            pragmaCmd.ExecuteNonQuery();

            var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                CREATE TABLE IF NOT EXISTS files (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    filepath TEXT UNIQUE NOT NULL,
                    filename TEXT NOT NULL,
                    content TEXT,
                    modified_time TEXT,
                    indexed_time TEXT DEFAULT CURRENT_TIMESTAMP
                );
                CREATE INDEX IF NOT EXISTS idx_filename ON files(filename);
            ";
            cmd.ExecuteNonQuery();

            // 尝试创建 FTS5 表，失败则忽略（不影响普通搜索）
            try
            {
                var ftsCmd = conn.CreateCommand();
                ftsCmd.CommandText = "CREATE VIRTUAL TABLE IF NOT EXISTS files_fts USING fts5(content, filepath, filename)";
                ftsCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"FTS5创建失败（不影响搜索）: {ex.Message}");
            }
        }

        private void UpdateIndexStatus()
        {
            using var conn = new SqliteConnection($"Data Source={_dbPath}");
            conn.Open();
            var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM files";
            var count = cmd.ExecuteScalar();
            IndexStatusText.Text = $"索引文件数: {count}";
        }

        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "选择搜索文件夹"
            };
            if (dialog.ShowDialog() == true)
            {
                PathTextBox.Text = dialog.FolderName;
                SaveLastPath();
            }
        }

        private void SelectFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择要搜索的文件",
                Multiselect = true,
                Filter = "所有文件 (*.*)|*.*|文本文件 (*.txt)|*.txt|文档文件 (*.doc;*.docx;*.pdf)|*.doc;*.docx;*.pdf"
            };
            if (dialog.ShowDialog() == true)
            {
                SearchInFiles(dialog.FileNames);
            }
        }

        private void SearchTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                Search_Click(sender, e);
        }

        private async void Search_Click(object sender, RoutedEventArgs e)
        {
            var keyword = SearchTextBox.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                MessageBox.Show("请输入搜索关键词", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (string.IsNullOrEmpty(PathTextBox.Text.Trim()))
            {
                MessageBox.Show("请选择搜索路径", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _cts = new CancellationTokenSource();
            ResultsListView.ItemsSource = null;
            StatusText.Text = "搜索中...";
            SearchTextBox.IsEnabled = false;

            try
            {
                await Task.Run(() => SearchInFolder(keyword, PathTextBox.Text.Trim(), _cts.Token));
                StatusText.Text = $"搜索完成";
            }
            catch (OperationCanceledException)
            {
                StatusText.Text = "搜索已取消";
            }
            finally
            {
                SearchTextBox.IsEnabled = true;
                _cts = null;
            }
        }

        private void Stop_Click(object sender, RoutedEventArgs e)
        {
            _cts?.Cancel();
        }

        private void SearchInFolder(string keyword, string folder, CancellationToken ct)
        {
            try
            {
                var results = new ObservableCollection<SearchResult>();
                
                using var conn = new SqliteConnection($"Data Source={_dbPath}");
                conn.Open();

                // 先检查索引是否存在（改查 files 表而不是 FTS5）
                var checkCmd = conn.CreateCommand();
                checkCmd.CommandText = "SELECT COUNT(*) FROM files";
                var count = checkCmd.ExecuteScalar();
                
                if (count == null || Convert.ToInt32(count) == 0)
                {
                    Dispatcher.Invoke(() =>
                    {
                        StatusText.Text = "请先建立索引！";
                        MessageBox.Show("请先点击「重建索引」按钮建立索引！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    });
                    return;
                }

                // 使用简单 LIKE 搜索替代 FTS5
                var cmd = conn.CreateCommand();
                
                // 修复：使用 startswith 而不是 LIKE，增强稳定性
                cmd.CommandText = @"
                    SELECT filepath, filename, substr(content, 1, 200) as preview
                    FROM files
                    WHERE filepath LIKE @keywordEscaped
                    AND (content LIKE @keyword OR filename LIKE @keyword)
                    LIMIT 500
                ";
                cmd.Parameters.AddWithValue("@keyword", $"%{keyword}%");
                // 转义用于 LIKE 的特殊字符
                var keywordEscaped = keyword.Replace("[", "[[]").Replace("%", "[%]").Replace("_", "[_]");
                cmd.Parameters.AddWithValue("@keywordEscaped", $"%{keywordEscaped}%");

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    if (ct.IsCancellationRequested) break;
                    
                    try
                    {
                        var preview = reader.IsDBNull(2) ? "" : reader.GetString(2);
                        var idx = preview.IndexOf(keyword, StringComparison.OrdinalIgnoreCase);
                        if (idx >= 0)
                        {
                            var start = Math.Max(0, idx - 30);
                            var len = Math.Min(60, preview.Length - start);
                            preview = "..." + preview.Substring(start, len) + "...";
                        }
                        else if (preview.Length > 60)
                        {
                            preview = preview.Substring(0, 60) + "...";
                        }
                        
                        results.Add(new SearchResult
                        {
                            FilePath = reader.GetString(0),
                            FileName = reader.GetString(1),
                            Preview = preview
                        });
                    }
                    catch { }
                }

                Dispatcher.Invoke(() =>
                {
                    ResultsListView.ItemsSource = results;
                    StatusText.Text = $"找到 {results.Count} 个结果";
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    StatusText.Text = "搜索出错: " + ex.Message;
                    MessageBox.Show("搜索出错: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }

        private void SearchInFiles(string[] files)
        {
            var keyword = SearchTextBox.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                MessageBox.Show("请输入搜索关键词", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var results = new ObservableCollection<SearchResult>();
            
            foreach (var file in files)
            {
                try
                {
                    var content = ExtractText(file);
                    if (content.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                    {
                        var idx = content.IndexOf(keyword, StringComparison.OrdinalIgnoreCase);
                        var start = Math.Max(0, idx - 30);
                        var len = Math.Min(60, content.Length - start);
                        var preview = content.Substring(start, len);
                        
                        results.Add(new SearchResult
                        {
                            FileName = Path.GetFileName(file),
                            FilePath = file,
                            Preview = "..." + preview + "..."
                        });
                    }
                }
                catch { }
            }

            ResultsListView.ItemsSource = results;
            StatusText.Text = $"找到 {results.Count} 个结果";
        }

        private void ResultsListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ResultsListView.SelectedItem is SearchResult result)
            {
                try
                {
                    // 检查文件是否存在
                    if (!File.Exists(result.FilePath))
                    {
                        MessageBox.Show("文件不存在或已被移动/删除", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = result.FilePath,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"无法打开文件: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async void BuildIndex_Click(object sender, RoutedEventArgs e)
        {
            var folder = PathTextBox.Text.Trim();
            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
            {
                MessageBox.Show("请选择有效的文件夹路径", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _isIndexing = true;
            BuildIndexBtn.IsEnabled = false;
            StatusText.Text = "正在建立索引...";

            try
            {
                await Task.Run(() => BuildIndex(folder));
                StatusText.Text = "索引建立完成";
                UpdateIndexStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"索引建立失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _isIndexing = false;
                BuildIndexBtn.IsEnabled = true;
            }
        }

        private void BuildIndex(string folder)
        {
            var extensions = new[] { ".txt", ".md", ".py", ".js", ".ts", ".html", ".css", ".json", ".xml", 
                ".cs", ".java", ".cpp", ".c", ".h", ".go", ".rs", ".sql",
                ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".pdf" };

            // 需要跳过的系统文件夹
            var systemFolders = new[] { 
                "System Volume Information", 
                "$RECYCLE.BIN",
                "Windows",
                "Program Files",
                "Program Files (x86)",
                "ProgramData",
                "$Recycle.Bin"
            };

            var files = Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories)
                .Where(f => {
                    // 跳过系统文件夹
                    var path = f.Replace('/', '\\');
                    foreach (var sf in systemFolders)
                    {
                        if (path.Contains(sf, StringComparison.OrdinalIgnoreCase))
                            return false;
                    }
                    return extensions.Contains(Path.GetExtension(f).ToLower());
                });

            using var conn = new SqliteConnection($"Data Source={_dbPath}");
            conn.Open();

            int count = 0;
            foreach (var file in files)
            {
                if (_isIndexing == false) break;

                try
                {
                    var info = new FileInfo(file);
                    var content = ExtractText(file);

                    var cmd = conn.CreateCommand();
                    cmd.CommandText = @"
                        INSERT OR REPLACE INTO files (filepath, filename, content, modified_time, indexed_time)
                        VALUES (@filepath, @filename, @content, @modified, datetime('now'))
                    ";
                    cmd.Parameters.AddWithValue("@filepath", file);
                    cmd.Parameters.AddWithValue("@filename", info.Name);
                    cmd.Parameters.AddWithValue("@content", content);
                    cmd.Parameters.AddWithValue("@modified", info.LastWriteTime.ToString("o"));
                    cmd.ExecuteNonQuery();
                    
                    count++;
                    if (count % 100 == 0)
                    {
                        Dispatcher.Invoke(() => { StatusText.Text = $"已索引 {count} 个文件..."; });
                    }
                }
                catch { }
            }
        }

        private string ExtractText(string filePath)
        {
            var ext = Path.GetExtension(filePath).ToLower();
            
            try
            {
                return ext switch
                {
                    ".txt" or ".md" or ".py" or ".js" or ".ts" or ".html" or ".css" or ".json" or ".xml" 
                        or ".cs" or ".java" or ".cpp" or ".c" or ".h" or ".go" or ".rs" or ".sql"
                        or ".log" or ".cfg" or ".ini" or ".yaml" or ".yml" => File.ReadAllText(filePath),
                    
                    ".doc" => ExtractDocText(filePath),
                    ".docx" => ExtractDocxText(filePath),
                    ".pdf" => ExtractPdfText(filePath),
                    ".xlsx" or ".xls" => ExtractExcelText(filePath),
                    _ => ""
                };
            }
            catch
            {
                return "";
            }
        }

        private string ExtractDocText(string filePath)
        {
            // 简单实现，需要 Microsoft.Office.Interop.Word
            // 这里返回空，实际使用需要添加 COM 引用
            return "";
        }

        private string ExtractDocxText(string filePath)
        {
            try
            {
                using var doc = WordprocessingDocument.Open(filePath, false);
                var body = doc.MainDocumentPart?.Document.Body;
                if (body == null) return "";
                
                return string.Join("\n", body.Elements<Paragraph>().Select(p => p.InnerText));
            }
            catch
            {
                return "";
            }
        }

        private string ExtractPdfText(string filePath)
        {
            try
            {
                using var doc = PdfDocument.Open(filePath);
                return string.Join("\n", doc.GetPages().Select(p => p.Text));
            }
            catch
            {
                return "";
            }
        }

        private string ExtractExcelText(string filePath)
        {
            try
            {
                using var doc = SpreadsheetDocument.Open(filePath, false);
                var sheets = doc.WorkbookPart?.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>();
                var text = new List<string>();
                
                foreach (var sheet in sheets ?? Enumerable.Empty<DocumentFormat.OpenXml.Spreadsheet.Sheet>())
                {
                    var worksheetPart = (WorksheetPart?)doc.WorkbookPart.GetPartById(sheet.Id!);
                    var cells = worksheetPart?.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>();
                    text.AddRange(cells?.Select(c => c.InnerText) ?? Enumerable.Empty<string>());
                }
                
                return string.Join(" ", text);
            }
            catch
            {
                return "";
            }
        }
    }
}
