using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace EtcOrder
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string? selectedFilePath;

        public MainWindow()
        {
            InitializeComponent();
            // 앱 아이콘 지정
            var icon = (System.Windows.Media.Imaging.BitmapImage)Application.Current.Resources["AppIcon"];
            this.Icon = icon;
        }

        private void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx";
            if (dlg.ShowDialog() == true)
            {
                selectedFilePath = dlg.FileName;
                txtSelectedFile.Text = selectedFilePath;
                btnConvert.IsEnabled = true;
                Log($"파일 선택됨: {selectedFilePath}");
            }
        }

        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFilePath) || !File.Exists(selectedFilePath))
            {
                Log("엑셀 파일을 먼저 선택하세요.");
                return;
            }
            try
            {
                string outputPath = Path.Combine(Path.GetDirectoryName(selectedFilePath)!, "거래처별_종합.xlsx");
                ExcelProcessor.Process(selectedFilePath, outputPath, Log);
                Log($"변환 완료! 결과 파일: {outputPath}");
            }
            catch (Exception ex)
            {
                Log($"오류 발생: {ex.Message}");
            }
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            txtLog.Text = string.Empty;
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && Path.GetExtension(files[0]).ToLower() == ".xlsx")
                {
                    selectedFilePath = files[0];
                    txtSelectedFile.Text = selectedFilePath;
                    btnConvert.IsEnabled = true;
                    Log($"파일 드롭됨: {selectedFilePath}");
                }
                else
                {
                    Log("엑셀(.xlsx) 파일만 지원합니다.");
                }
            }
        }

        private void Log(string msg)
        {
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}\n");
            txtLog.ScrollToEnd();
        }
    }
}