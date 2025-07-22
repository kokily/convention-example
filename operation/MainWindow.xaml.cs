using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using OfficeOpenXml;

namespace IntegratedApp
{
    public partial class MainWindow : Window
    {
        private string? selectedFilePath;
        private AppType currentAppType = AppType.SamsungOrder;

        public enum AppType
        {
            EtcOrder,
            SamsungOrder,
            SamsungStatements,
            Expendables
        }

        public MainWindow()
        {
            InitializeComponent();
            
            // 초기 라디오 버튼 설정 (이벤트 발생 방지)
            rbSamsungOrder.IsChecked = true;
            
            // 초기 상태 설정
            UpdateAppStatus();
        }

        private void AppSelection_Changed(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton radioButton)
            {
                // 선택된 앱 타입 결정
                if (radioButton == rbEtcOrder)
                    currentAppType = AppType.EtcOrder;
                else if (radioButton == rbSamsungOrder)
                    currentAppType = AppType.SamsungOrder;
                else if (radioButton == rbSamsungStatements)
                    currentAppType = AppType.SamsungStatements;
                else if (radioButton == rbExpendables)
                    currentAppType = AppType.Expendables;

                // UI 컨트롤들이 초기화된 후에만 실행
                if (txtSelectedFile != null && btnConvert != null && txtLog != null)
                {
                    // 앱 변경 시 초기화
                    InitializeApp();
                    UpdateAppStatus();
                }
            }
        }

                            private void InitializeApp()
                    {
                        // 선택된 파일 초기화
                        selectedFilePath = null;
                        txtSelectedFile.Text = "";
                        btnConvert.IsEnabled = false;

                        // 로그 초기화
                        txtLog.Text = "";
                        
                        // 드롭 상태 초기화
                        UpdateDropStatus("엑셀 파일을 여기에 드래그하거나 아래 버튼을 클릭하세요");
                    }

        private void UpdateAppStatus()
        {
            string appName = currentAppType switch
            {
                AppType.EtcOrder => "거래처별 엑셀 변환기",
                AppType.SamsungOrder => "삼성웰스토리 발주서 전처리기",
                AppType.SamsungStatements => "삼성웰스토리 결산서 전처리기",
                AppType.Expendables => "컨벤션 소모품 결산",
                _ => "알 수 없는 앱"
            };

            Log($"=== {appName} 모드로 전환되었습니다 ===");
        }



        private void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            
            // 앱 타입에 따른 파일 필터 설정
            dlg.Filter = currentAppType switch
            {
                AppType.SamsungStatements => "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls",
                _ => "Excel Files (*.xlsx)|*.xlsx"
            };

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
                // 앱 타입에 따른 처리 로직 실행
                switch (currentAppType)
                {
                    case AppType.EtcOrder:
                        ProcessEtcOrder();
                        break;
                    case AppType.SamsungOrder:
                        ProcessSamsungOrder();
                        break;
                    case AppType.SamsungStatements:
                        ProcessSamsungStatements();
                        break;
                    case AppType.Expendables:
                        ProcessExpendables();
                        break;
                }
            }
            catch (Exception ex)
            {
                Log($"오류 발생: {ex.Message}");
            }
        }

        private void ProcessEtcOrder()
        {
            Log("=== 거래처별 엑셀 변환기 처리 시작 ===");
            string outputPath = Path.Combine(Path.GetDirectoryName(selectedFilePath)!, "거래처별_종합.xlsx");
            
            try
            {
                EtcOrderProcessor.Process(selectedFilePath, outputPath, Log);
                Log($"변환 완료! 결과 파일: {outputPath}");
            }
            catch (Exception ex)
            {
                Log($"오류 발생: {ex.Message}");
            }
        }

        private void ProcessSamsungOrder()
        {
            Log("=== 삼성웰스토리 발주서 전처리기 처리 시작 ===");
            string outputPath = Path.Combine(Path.GetDirectoryName(selectedFilePath)!, "삼성웰스토리_발주서_결과.xlsx");
            
            try
            {
                SamsungOrderProcessor.Process(selectedFilePath, outputPath, Log);
                Log($"처리 완료! 결과 파일: {outputPath}");
            }
            catch (Exception ex)
            {
                Log($"오류 발생: {ex.Message}");
            }
        }

        private void ProcessSamsungStatements()
        {
            Log("=== 삼성웰스토리 결산서 전처리기 처리 시작 ===");
            string outputPath = Path.Combine(Path.GetDirectoryName(selectedFilePath)!, "삼성웰스토리_결산서_결과.xlsx");
            
            try
            {
                SamsungStatementsProcessor.Process(selectedFilePath, outputPath, Log);
                Log($"처리 완료! 결과 파일: {outputPath}");
            }
            catch (Exception ex)
            {
                Log($"오류 발생: {ex.Message}");
            }
        }

        private void ProcessExpendables()
        {
            Log("=== 컨벤션 소모품 결산 처리 시작 ===");
            string outputPath = Path.Combine(Path.GetDirectoryName(selectedFilePath)!, "컨벤션_소모품_결산_결과.xlsx");
            
            try
            {
                ExpendablesProcessor.Process(selectedFilePath, outputPath, Log);
                Log($"결산 완료! 결과 파일: {outputPath}");
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



                            private void Log(string msg)
                    {
                        txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}\n");
                        txtLog.ScrollToEnd();
                    }

                    private void Window_DragEnter(object sender, DragEventArgs e)
                    {
                        if (e.Data.GetDataPresent(DataFormats.FileDrop))
                        {
                            e.Effects = DragDropEffects.Copy;
                            UpdateDropStatus("파일을 놓으세요");
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
                                string extension = Path.GetExtension(filePath).ToLower();
                                
                                // 엑셀 파일인지 확인
                                if (extension == ".xlsx" || extension == ".xls")
                                {
                                    selectedFilePath = filePath;
                                    txtSelectedFile.Text = selectedFilePath;
                                    btnConvert.IsEnabled = true;
                                    Log($"파일 드롭됨: {selectedFilePath}");
                                    UpdateDropStatus("파일이 성공적으로 로드되었습니다");
                                }
                                else
                                {
                                    Log("엑셀 파일(.xlsx, .xls)만 지원됩니다.");
                                    UpdateDropStatus("엑셀 파일만 지원됩니다");
                                }
                            }
                        }
                        e.Handled = true;
                    }

                    private void UpdateDropStatus(string message)
                    {
                        if (txtDropStatus != null)
                        {
                            txtDropStatus.Text = message;
                        }
                    }
    }
} 