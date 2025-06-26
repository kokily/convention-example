using Microsoft.Win32;
using System.Windows;
using System.Windows.Threading;
using System.IO;

namespace SamsungStatements
{
    public partial class MainWindow : Window
    {
        private string? selectedFilePath;
        private readonly LogManager logManager;

        public MainWindow()
        {
            InitializeComponent();
            logManager = new LogManager(txtLog);
            LogMessage("삼성웰스토리 결산서 전처리기가 시작되었습니다.");
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string filePath = files[0];
                    string ext = Path.GetExtension(filePath).ToLower();

                    if (ext != ".xlsx" && ext != ".xls")
                    {
                        MessageBox.Show($"지원하지 않는 파일 형식: {ext}\n.xlsx 또는 .xls 파일을 드롭해 주세요", 
                                       "파일 형식 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
                        txtDropStatus.Text = "잘못된 파일 형식입니다. .xlsx 또는 .xls 파일을 드롭해주세요";
                        selectedFilePath = null;
                        btnConvert.IsEnabled = false;
                        return;
                    }

                    selectedFilePath = filePath;
                    txtSelectedFile.Text = $"선택된 파일: {Path.GetFileName(selectedFilePath)}";
                    txtDropStatus.Text = $"파일 드롭됨: {Path.GetFileName(filePath)}\n'엑셀 처리 시작' 버튼을 클릭하세요";
                    btnConvert.IsEnabled = true;
                    LogMessage($"파일이 드롭되었습니다: {selectedFilePath}");
                }
            }
            e.Handled = true;
        }

        private void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "엑셀 파일 선택",
                Filter = "Excel 파일 (*.xlsx;*.xls)|*.xlsx;*.xls|모든 파일 (*.*)|*.*",
                FilterIndex = 1
            };

            if (openFileDialog.ShowDialog() == true)
            {
                selectedFilePath = openFileDialog.FileName;
                txtSelectedFile.Text = $"선택된 파일: {Path.GetFileName(selectedFilePath)}";
                txtDropStatus.Text = $"파일 선택됨: {Path.GetFileName(selectedFilePath)}\n'엑셀 처리 시작' 버튼을 클릭하세요";
                btnConvert.IsEnabled = true;
                LogMessage($"파일이 선택되었습니다: {selectedFilePath}");
            }
        }

        private async void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFilePath))
            {
                MessageBox.Show("먼저 엑셀 파일을 선택하거나 드롭해주세요.", "파일 없음", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            btnConvert.IsEnabled = false;
            btnSelectFile.IsEnabled = false;
            txtDropStatus.Text = $"파일 처리 중: {Path.GetFileName(selectedFilePath)}...";

            try
            {
                LogMessage("변환 작업을 시작합니다...");
                
                // 비동기로 전체 변환 프로세스 실행
                await Task.Run(() =>
                {
                    ProcessExcel.ProcessExcelFile(selectedFilePath, logManager);
                });

                string outputFileName = "결산서.xlsx";
                string outputFilePath = Path.Combine(Path.GetDirectoryName(selectedFilePath) ?? "", outputFileName);

                LogMessage("전체 변환 프로세스가 성공적으로 완료되었습니다!");
                LogMessage($"결과 파일: {outputFilePath}");

                MessageBox.Show($"Excel 파일이 성공적으로 처리되어\n{outputFilePath} (으)로 저장되었습니다.", 
                               "성공", MessageBoxButton.OK, MessageBoxImage.Information);

                txtDropStatus.Text = "처리 완료! 다른 Excel 파일을 드롭하거나 선택하세요.";
                selectedFilePath = null;
                btnConvert.IsEnabled = false;
            }
            catch (Exception ex)
            {
                LogMessage($"오류 발생: {ex.Message}");
                MessageBox.Show($"파일 처리 실패: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                txtDropStatus.Text = "파일 처리 중 오류 발생. 다시 드롭하거나 선택하세요.";
                selectedFilePath = null;
                btnConvert.IsEnabled = false;
            }
            finally
            {
                btnSelectFile.IsEnabled = true;
            }
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            txtLog.Clear();
            LogMessage("로그가 지워졌습니다.");
        }

        private void LogMessage(string message)
        {
            logManager.LogMessage(message);
        }
    }
} 