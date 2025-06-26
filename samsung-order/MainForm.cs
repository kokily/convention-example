using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace SamsungOrder
{
    public partial class MainForm : Form
    {
        private TableLayoutPanel tableLayout = null!;
        private TextBox logTextBox = null!;
        private Label statusLabel = null!;
        private Button fileButton = null!;
        private Button processButton = null!;
        private Button clearButton = null!;
        private string droppedFilePath = "";

        public MainForm()
        {
            InitializeComponent();
            SetupUI();
            SetupConsoleRedirect();
        }

        private void InitializeComponent()
        {
            this.Text = "삼성웰스토리 결산서 전처리기";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AllowDrop = true;
            this.DragEnter += MainForm_DragEnter;
            this.DragDrop += MainForm_DragDrop;
        }

        private void SetupUI()
        {
            // 전체 레이아웃
            tableLayout = new TableLayoutPanel
            {
                RowCount = 3,
                ColumnCount = 1,
                Dock = DockStyle.Fill,
            };
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 90F)); // 안내 영역
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 로그
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 70F)); // 버튼
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // 안내 Panel
            var guidePanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(220, 245, 255),
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(10, 10, 10, 10)
            };
            var guideLabel1 = new Label
            {
                Text = "여기에 엑셀파일을 드래그 앤 드롭하세요",
                Dock = DockStyle.Top,
                Font = new Font("맑은 고딕", 14, FontStyle.Bold),
                ForeColor = Color.RoyalBlue,
                TextAlign = ContentAlignment.MiddleCenter,
                Height = 32
            };
            var guideLabel2 = new Label
            {
                Text = "또는 아래 버튼을 클릭하여 파일을 선택하세요",
                Dock = DockStyle.Top,
                Font = new Font("맑은 고딕", 10, FontStyle.Regular),
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleCenter,
                Height = 24
            };
            guidePanel.Controls.Add(guideLabel2);
            guidePanel.Controls.Add(guideLabel1);

            // 로그 TextBox
            logTextBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 11),
                BackColor = Color.White,
                ForeColor = Color.Black,
                WordWrap = false
            };

            // 버튼 패널 (왼쪽 정렬)
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = false,
                Height = 50,
                Padding = new Padding(10, 10, 0, 10)
            };
            fileButton = new Button { Text = "파일 선택", AutoSize = true, MinimumSize = new Size(110, 36), Font = new Font("맑은 고딕", 10) };
            processButton = new Button { Text = "엑셀 처리 시작", AutoSize = true, MinimumSize = new Size(110, 36), Font = new Font("맑은 고딕", 10), Enabled = false };
            clearButton = new Button { Text = "로그 지우기", AutoSize = true, MinimumSize = new Size(110, 36), Font = new Font("맑은 고딕", 10) };
            fileButton.Click += FileButton_Click;
            processButton.Click += ProcessButton_Click;
            clearButton.Click += (s, e) => logTextBox.Clear();
            fileButton.Margin = new Padding(0, 0, 10, 0);
            processButton.Margin = new Padding(0, 0, 10, 0);
            clearButton.Margin = new Padding(0, 0, 0, 0);
            buttonPanel.Controls.Add(fileButton);
            buttonPanel.Controls.Add(processButton);
            buttonPanel.Controls.Add(clearButton);

            // 상태 라벨 (로그 위에, 내부적으로만 사용)
            statusLabel = new Label
            {
                Text = "",
                Dock = DockStyle.Top,
                Font = new Font("맑은 고딕", 9, FontStyle.Regular),
                ForeColor = Color.Black,
                Height = 0,
                Visible = false
            };

            // 레이아웃에 추가
            tableLayout.Controls.Add(guidePanel, 0, 0);
            tableLayout.Controls.Add(logTextBox, 0, 1);
            tableLayout.Controls.Add(buttonPanel, 0, 2);

            this.Controls.Add(tableLayout);
        }

        private void SetupConsoleRedirect()
        {
            var writer = new TextBoxWriter(logTextBox);
            Console.SetOut(writer);
        }

        private void FileButton_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files (*.xlsx)|*.xlsx";
                ofd.Title = "엑셀 파일 선택";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    droppedFilePath = ofd.FileName;
                    processButton.Enabled = true;
                    LogWithTime($"파일 선택됨: {Path.GetFileName(droppedFilePath)}");
                }
            }
        }

        private void MainForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string filePath = files[0];
                    string ext = Path.GetExtension(filePath).ToLower();

                    if (ext != ".xlsx")
                    {
                        MessageBox.Show($"지원하지 않는 파일 형식: {ext}.\n.xlsx 파일을 드랍해주세요", 
                            "잘못된 파일 형식", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        LogWithTime("잘못된 파일 형식입니다. .xlsx 파일을 드랍하세요");
                        droppedFilePath = "";
                        processButton.Enabled = false;
                        return;
                    }

                    droppedFilePath = filePath;
                    processButton.Enabled = true;
                    LogWithTime($"파일 드랍됨: {Path.GetFileName(filePath)}");
                }
            }
        }

        private async void ProcessButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(droppedFilePath))
            {
                MessageBox.Show("먼저 엑셀 파일을 선택하거나 드랍하세요", "파일 없음", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            LogWithTime($"엑셀 처리 시작: {Path.GetFileName(droppedFilePath)}");
            processButton.Enabled = false;

            try
            {
                await ProcessExcelFileAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"파일 처리 중 오류 발생: {ex.Message}", "오류", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogWithTime("파일 처리 중 오류 발생. 다시 시도하세요");
                droppedFilePath = "";
            }
            finally
            {
                processButton.Enabled = true;
            }
        }

        private async Task ProcessExcelFileAsync()
        {
            await Task.Run(() =>
            {
                string inputFileName = Path.GetFileName(droppedFilePath);

                List<TransData> transData;
                try
                {
                    transData = Extractor.ExtractRawDataFromExcel(droppedFilePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ERROR: ExtractRawDataFromExcel 추출 오류: {ex.Message}");
                    throw new Exception($"파일처리 실패 (데이터 추출): {ex.Message}", ex);
                }

                Console.WriteLine($"디버그: RawData 추출 및 변환 완료. 총 TransData 행 수: {transData.Count}");

                if (transData.Count == 0)
                {
                    throw new Exception("원본 파일에서 읽을 데이터가 없습니다.");
                }

                List<TransData> resultData;
                try
                {
                    resultData = Utils.TransformData(transData);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ERROR: 데이터 가공 오류: {ex.Message}");
                    throw new Exception($"데이터 가공 실패 (데이터 변환): {ex.Message}", ex);
                }

                Console.WriteLine($"디버그: 데이터 변환 완료. 총 ResultData 행 수: {resultData.Count}");

                try
                {
                    ExcelWriter.WriteProcessedDataToExcel(resultData, droppedFilePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ERROR: 엑셀 파일 저장 오류: {ex.Message}");
                    throw new Exception($"파일 저장 실패: {ex.Message}", ex);
                }

                Console.WriteLine("디버그: 엑셀 파일 저장 완료");

                this.Invoke(() =>
                {
                    string outputFileName = $"분류된_{Path.GetFileNameWithoutExtension(droppedFilePath)}.xlsx";
                    string outputFilePath = Path.Combine(Path.GetDirectoryName(droppedFilePath)!, outputFileName);
                    MessageBox.Show($"엑셀 파일이 성공적으로 처리되어\n{outputFilePath} (으)로 저장되었습니다", 
                        "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LogWithTime($"엑셀 파일 저장 완료: {outputFilePath}");
                });

                droppedFilePath = "";
            });
        }

        private void LogWithTime(string message)
        {
            string time = DateTime.Now.ToString("[yyyy-MM-dd HH:mm:ss]");
            logTextBox.AppendText($"{time} {message}{Environment.NewLine}");
            logTextBox.SelectionStart = logTextBox.Text.Length;
            logTextBox.ScrollToCaret();
        }
    }

    // Console.WriteLine을 TextBox로 리다이렉트하는 클래스
    public class TextBoxWriter : System.IO.TextWriter
    {
        private TextBox textBox;

        public TextBoxWriter(TextBox textBox)
        {
            this.textBox = textBox;
        }

        public override System.Text.Encoding Encoding => System.Text.Encoding.UTF8;

        public override void Write(char value)
        {
            if (textBox.InvokeRequired)
            {
                textBox.Invoke(new Action(() => Write(value)));
                return;
            }

            textBox.AppendText(value.ToString());
            textBox.SelectionStart = textBox.Text.Length;
            textBox.ScrollToCaret();
        }

        public override void Write(string? value)
        {
            if (textBox.InvokeRequired)
            {
                textBox.Invoke(new Action(() => Write(value)));
                return;
            }

            if (value != null)
            {
                textBox.AppendText(value);
                textBox.SelectionStart = textBox.Text.Length;
                textBox.ScrollToCaret();
            }
        }

        public override void WriteLine(string? value)
        {
            if (value != null)
                textBox.AppendText(value + Environment.NewLine);
            else
                textBox.AppendText(Environment.NewLine);
            textBox.SelectionStart = textBox.Text.Length;
            textBox.ScrollToCaret();
        }
    }
} 