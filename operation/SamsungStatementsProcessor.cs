using OfficeOpenXml;
using System.Text;
using System.IO;
using System.Globalization;
using System.Linq;

namespace IntegratedApp
{
    public static class SamsungStatementsProcessor
    {
        public static void Process(string inputPath, string outputPath, Action<string>? log = null)
        {
            log?.Invoke("삼성웰스토리 결산서 전처리기가 시작되었습니다.");
            
            // EPPlus 라이센스 설정 (가장 먼저 설정)
            ExcelPackage.License.SetNonCommercialPersonal("김현성");
            
            try
            {
                log?.Invoke($"파일이 선택되었습니다: {inputPath}");
                log?.Invoke("변환 작업을 시작합니다...");

                // 1단계: CsvToJson - 엑셀 파일을 RawData로 변환
                var rawData = CsvToJson.ConvertCsvToJson(inputPath, log);
                log?.Invoke($"DEBUG: CsvToJson 완료 ===> 로드 RawData 행: {rawData.Count}");

                if (rawData.Count == 0)
                {
                    throw new InvalidOperationException("원본 엑셀파일의 데이터가 부족합니다.");
                }

                // 2단계: ManufactureJson - RawData를 TransData로 변환
                var inputFileName = Path.GetFileNameWithoutExtension(inputPath);
                var transData = ManufactureJson.ManufactureJsonData(rawData, inputFileName, log);
                log?.Invoke($"DEBUG: ManufactureJson 완료 ===> 변환된 TransData 수: {transData.Count}");

                if (transData.Count == 0)
                {
                    throw new InvalidOperationException("데이터 변환 후 유효한 데이터가 부족합니다.");
                }

                // 3단계: 새로운 Excel 워크북 생성
                using var workBook = new ExcelPackage();

                // 기본 Sheet1 삭제 시도
                try
                {
                    var defaultSheet = workBook.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Sheet1");
                    if (defaultSheet != null)
                    {
                        workBook.Workbook.Worksheets.Delete(defaultSheet);
                        log?.Invoke("DEBUG: 기본 Sheet1 삭제 완료");
                    }
                }
                catch (Exception ex)
                {
                    log?.Invoke($"WARN: Sheet1 삭제 실패 (정상동작일 수 있음): {ex.Message}");
                }

                // 4단계: 데이터 분류 및 시트 추가
                ClassificationItems.ClassifyItems(transData, workBook, outputPath, log);

                log?.Invoke("전체 변환 프로세스가 성공적으로 완료되었습니다!");
                log?.Invoke($"결과 파일: {outputPath}");
            }
            catch (Exception ex)
            {
                log?.Invoke($"ERROR: ProcessExcelFile 오류: {ex.Message}");
                throw new InvalidOperationException($"ProcessExcelFile 오류: {ex.Message}", ex);
            }
        }

        public static class CsvToJson
        {
            public static List<RawData> ConvertCsvToJson(string filePath, Action<string>? log = null)
            {
                log?.Invoke($"DEBUG: CsvToJson 시작 =====================> 파일: {filePath}");

                var jsonData = new List<RawData>();

                // EPPlus 라이센스 설정
                ExcelPackage.License.SetNonCommercialPersonal("김현성");

                using var package = new ExcelPackage(new FileInfo(filePath));
                var worksheet = package.Workbook.Worksheets[0]; // 첫 번째 시트

                if (worksheet == null)
                {
                    log?.Invoke("ERROR: 엑셀 파일에 시트가 없습니다");
                    throw new InvalidOperationException("엑셀 파일에 시트가 없습니다");
                }

                log?.Invoke($"DEBUG: 시트 이름 {worksheet.Name}");

                // 전체 데이터 범위 가져오기
                var dataRange = worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column, 
                                               worksheet.Dimension.End.Row, worksheet.Dimension.End.Column];
                
                var rows = dataRange.Value as object[,];
                
                if (rows == null)
                {
                    log?.Invoke("ERROR: 엑셀 파일에서 데이터를 읽을 수 없습니다.");
                    throw new InvalidOperationException("엑셀 파일에서 데이터를 읽을 수 없습니다.");
                }

                int rowCount = rows.GetLength(0);
                int colCount = rows.GetLength(1);

                log?.Invoke($"DEBUG: 읽은 총 행 수 (헤더포함) {rowCount}");

                if (rowCount < 2)
                {
                    log?.Invoke("ERROR: 엑셀 파일에 데이터가 충분하지 않습니다.");
                    throw new InvalidOperationException("엑셀 파일에 데이터가 충분하지 않습니다.");
                }

                for (int rIdx = 0; rIdx < rowCount; rIdx++)
                {
                    var aColVal = GetCellValue(rows, rIdx, 0)?.ToString()?.Trim() ?? "";

                    if (aColVal == "기간별구매현황[일자별]" || aColVal == "순번")
                    {
                        log?.Invoke($"DEBUG: 제목 라인 스킵 (행 {rIdx + 1})");
                        continue;
                    }

                    var data = new RawData
                    {
                        순번 = GetCellValue(rows, rIdx, 0)?.ToString() ?? "",
                        입고일자 = GetCellValue(rows, rIdx, 1)?.ToString() ?? "",
                        품번 = GetCellValue(rows, rIdx, 2)?.ToString() ?? "",
                        품명 = GetCellValue(rows, rIdx, 3)?.ToString() ?? "",
                        규격 = GetCellValue(rows, rIdx, 4)?.ToString() ?? "",
                        수량 = GetCellValue(rows, rIdx, 5)?.ToString() ?? "",
                        단위 = GetCellValue(rows, rIdx, 6)?.ToString() ?? "",
                        단가 = GetCellValue(rows, rIdx, 7)?.ToString() ?? "",
                        금액 = GetCellValue(rows, rIdx, 8)?.ToString() ?? "",
                        부가세 = GetCellValue(rows, rIdx, 9)?.ToString() ?? "",
                        합계금액 = GetCellValue(rows, rIdx, 10)?.ToString() ?? "",
                        거래처명 = GetCellValue(rows, rIdx, 11)?.ToString() ?? "",
                        적요 = GetCellValue(rows, rIdx, 12)?.ToString() ?? "",
                        특이사항 = GetCellValue(rows, rIdx, 13)?.ToString() ?? "",
                        현장명 = GetCellValue(rows, rIdx, 14)?.ToString() ?? "",
                        PJT코드 = GetCellValue(rows, rIdx, 15)?.ToString() ?? "",
                        PJT명 = GetCellValue(rows, rIdx, 16)?.ToString() ?? "",
                        입고창고 = GetCellValue(rows, rIdx, 17)?.ToString() ?? ""
                    };

                    jsonData.Add(data);
                    
                    // 100개 단위로 진행률 표시
                    if ((rIdx + 1) % 100 == 0)
                    {
                        log?.Invoke($"DEBUG: 엑셀 데이터 읽기 진행률: {rIdx + 1}/{rowCount} 행 처리 완료");
                    }
                }

                log?.Invoke($"DEBUG: CsvToJson 종료. 변환된 데이터 수량 : {jsonData.Count}");
                log?.Invoke($"DEBUG: CsvToJson 완료 ===> 로드 RawData 행: {jsonData.Count}");
                return jsonData;
            }

            private static object? GetCellValue(object[,] rows, int rowIndex, int colIndex)
            {
                try
                {
                    if (rowIndex < rows.GetLength(0) && colIndex < rows.GetLength(1))
                    {
                        return rows[rowIndex, colIndex];
                    }
                    return null;
                }
                catch
                {
                    return null;
                }
            }
        }

        public static class ManufactureJson
        {
            public static List<TransData> ManufactureJsonData(List<RawData> rawData, string inputFileName, Action<string>? log = null)
            {
                log?.Invoke($"DEBUG: ManufactureJson 시작 =====================> 입력파일명: {inputFileName}");

                var transDataList = new List<TransData>();
                int processedCount = 0;

                foreach (var raw in rawData)
                {
                    // 필수 필드 검증
                    if (string.IsNullOrWhiteSpace(raw.품명))
                    {
                        log?.Invoke($"DEBUG: 품명이 비어있는 행 스킵: {raw.순번}");
                        continue;
                    }

                    // 디버그 로그
                    log?.Invoke($"DEBUG: 변환 전 업체명='{raw.거래처명}', 현장명='{raw.현장명}'");

                    // 날짜 변환
                    string purchaseDate = ConvertDate(raw.입고일자);
                    
                    // 수량 변환
                    double quantity = ParseDouble(raw.수량);
                    
                    // 단가 변환
                    double unitPrice = ParseDouble(raw.단가);
                    if (string.IsNullOrWhiteSpace(raw.단가))
                    {
                        log?.Invoke($"WARN: 단가 파싱 오류 (행 {raw.순번}): 값: '{raw.단가}', 0으로 처리");
                    }
                    
                    // 금액 변환
                    double amount = ParseDouble(raw.금액);

                    // 입고일자 파싱 오류 체크
                    if (raw.입고일자 == "- -")
                    {
                        log?.Invoke($"WARN: 입고일자 파싱 오류 (행 {raw.순번}): 값: '{raw.입고일자}', 0으로 처리");
                    }

                    // 부가세 값 파싱
                    double vatAmount = ParseDouble(raw.부가세);
                    
                    var transData = new TransData
                    {
                        구매일자 = purchaseDate,
                        납품장소 = raw.현장명,
                        품명 = raw.품명,
                        규격 = raw.규격,
                        세 = vatAmount == 0 ? "면세" : "과세",
                        단위 = raw.단위,
                        수량 = quantity,
                        단가 = unitPrice,
                        금액 = amount,
                        부가세 = vatAmount,
                        업체명 = raw.거래처명
                    };

                    transDataList.Add(transData);
                    processedCount++;

                    // 100개 단위로 진행률 표시
                    if (processedCount % 100 == 0)
                    {
                        log?.Invoke($"DEBUG: 데이터 변환 진행률: {processedCount}/{rawData.Count} 건 처리 완료");
                    }
                }

                // 통계 정보 출력
                var supplierCount = transDataList.Select(x => x.업체명).Distinct().Count();
                var itemCount = transDataList.Select(x => x.품명).Distinct().Count();
                var totalAmount = transDataList.Sum(x => x.금액);
                var taxFreeCount = transDataList.Count(x => x.부가세 == 0);
                var taxFreeAmount = transDataList.Where(x => x.부가세 == 0).Sum(x => x.금액);
                var taxableCount = transDataList.Count(x => x.부가세 > 0);
                var taxableAmount = transDataList.Where(x => x.부가세 > 0).Sum(x => x.금액 * 1.1); // 과세는 부가세 포함이므로 1.1 곱함

                log?.Invoke($"DEBUG: 통계 - 업체 수: {supplierCount}, 품목 수: {itemCount}, 총 금액: {totalAmount:N0}원");
                log?.Invoke($"DEBUG: 면세 통계 - 건수: {taxFreeCount}, 금액: {taxFreeAmount:N0}원");
                log?.Invoke($"DEBUG: 과세 통계 - 건수: {taxableCount}, 금액: {taxableAmount:N0}원");

                log?.Invoke($"DEBUG: ManufactureJson 종료. 변환된 데이터 수량 : {transDataList.Count}");
                log?.Invoke($"DEBUG: ManufactureJson 완료 ===> 변환된 TransData 수: {transDataList.Count}");
                return transDataList;
            }

            private static string ConvertDate(string dateStr)
            {
                if (string.IsNullOrWhiteSpace(dateStr) || dateStr == "- -")
                    return "";

                // 다양한 날짜 형식 처리
                if (DateTime.TryParse(dateStr, out var dt))
                    return dt.ToString("yyyy-MM-dd");
                
                if (DateTime.TryParseExact(dateStr, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt2))
                    return dt2.ToString("yyyy-MM-dd");
                
                if (double.TryParse(dateStr, out var serial))
                {
                    var baseDate = new DateTime(1899, 12, 30);
                    var excelDate = baseDate.AddDays(serial);
                    return excelDate.ToString("yyyy-MM-dd");
                }

                return dateStr;
            }

            private static double ParseDouble(string value)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return 0;

                // 쉼표 제거
                value = value.Replace(",", "");
                
                if (double.TryParse(value, out double result))
                    return result;
                
                return 0;
            }

            private static string ParseTaxInfo(string taxStr)
            {
                if (string.IsNullOrWhiteSpace(taxStr))
                    return "면세";

                // 부가세가 0이거나 비어있으면 면세
                if (double.TryParse(taxStr.Replace(",", ""), out double taxAmount))
                {
                    return taxAmount == 0 ? "면세" : "과세";
                }

                return "면세";
            }
        }

        public static class ClassificationItems
        {
            public static void ClassifyItems(List<TransData> transData, ExcelPackage workBook, string outputPath, Action<string>? log = null)
            {
                log?.Invoke($"DEBUG: ClassificationItems 시작 ===================> TransData 개수: {transData.Count}");

                // 1차 분류: 업체명별 (로그에서 보이는 실제 분류)
                var contractData = transData.Where(x => x.업체명 == "삼성웰스토리").ToList();
                var nonContractData = transData.Where(x => x.업체명 == "삼성웰스토리(비)").ToList();
                
                log?.Invoke($"DEBUG: 1차 분류 업체명별. 계약: {contractData.Count}, 비계약: {nonContractData.Count}");

                // 2차 분류: 현장명별 (로그에서 보이는 실제 분류)
                var contractItems = contractData.Where(x => !x.납품장소.Contains("직원")).ToList();
                var contractEmployeeItems = contractData.Where(x => x.납품장소.Contains("직원")).ToList();
                var nonContractItems = nonContractData.Where(x => !x.납품장소.Contains("직원")).ToList();
                var nonContractEmployeeItems = nonContractData.Where(x => x.납품장소.Contains("직원")).ToList();

                log?.Invoke($"DEBUG: 2차 분류 계약: {contractItems.Count}, 비계약: {nonContractItems.Count}, 직원: {contractEmployeeItems.Count}, 직원 비계약: {nonContractEmployeeItems.Count}");

                // 계약 식자재 시트 생성 (로그에서 보이는 실제 시트명)
                if (contractItems.Count > 0)
                {
                    var contractSheet = workBook.Workbook.Worksheets.Add("계약 식자재");
                    CreateStatementSheet(contractSheet, contractItems, log);
                    log?.Invoke($"DEBUG: '계약 식자재' 시트에 {contractItems.Count}개 데이터 저장 완료");
                }

                // 계약 직원 식자재 시트 생성 (로그에서 보이는 실제 시트명)
                if (contractEmployeeItems.Count > 0)
                {
                    var employeeSheet = workBook.Workbook.Worksheets.Add("계약 직원 식자재");
                    CreateStatementSheet(employeeSheet, contractEmployeeItems, log);
                    log?.Invoke($"DEBUG: '계약 직원 식자재' 시트에 {contractEmployeeItems.Count}개 데이터 저장 완료");
                }

                // 비계약 식자재 시트 생성
                if (nonContractItems.Count > 0)
                {
                    var nonContractSheet = workBook.Workbook.Worksheets.Add("비계약 식자재");
                    CreateStatementSheet(nonContractSheet, nonContractItems, log);
                    log?.Invoke($"DEBUG: '비계약 식자재' 시트에 {nonContractItems.Count}개 데이터 저장 완료");
                }

                // 비계약 직원 식자재 시트 생성
                if (nonContractEmployeeItems.Count > 0)
                {
                    var nonContractEmployeeSheet = workBook.Workbook.Worksheets.Add("비계약 직원 식자재");
                    CreateStatementSheet(nonContractEmployeeSheet, nonContractEmployeeItems, log);
                    log?.Invoke($"DEBUG: '비계약 직원 식자재' 시트에 {nonContractEmployeeItems.Count}개 데이터 저장 완료");
                }

                // 종합 시트 생성 (이미지형)
                var summarySheet = workBook.Workbook.Worksheets.Add("종합");
                CreateSummarySheet(summarySheet, transData, log);
                log?.Invoke($"DEBUG: '종합' 시트(이미지형) 추가 완료");

                // 파일 저장
                log?.Invoke($"DEBUG: 엑셀 파일 저장 시도: {outputPath}");
                workBook.SaveAs(new FileInfo(outputPath));
                log?.Invoke($"DEBUG: 엑셀파일 '{outputPath}' 저장 완료");
                log?.Invoke("DEBUG: 분류 시트들 추가 및 Excel 파일 저장 완료");
            }

            private static void CreateStatementSheet(ExcelWorksheet sheet, List<TransData> data, Action<string>? log = null)
            {
                // 헤더 설정
                string[] headers = { "순번", "구매일자", "납품장소", "품명", "규격", "세", "단위", "수량", "단가", "금액", "업체명" };
                
                for (int i = 0; i < headers.Length; i++)
                {
                    sheet.Cells[1, i + 1].Value = headers[i];
                }

                // 헤더 스타일
                using (var range = sheet.Cells[1, 1, 1, headers.Length])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                }

                // 데이터 추가
                for (int i = 0; i < data.Count; i++)
                {
                    var item = data[i];
                    int row = i + 2;

                    sheet.Cells[row, 1].Value = i + 1;
                    sheet.Cells[row, 2].Value = item.구매일자;
                    sheet.Cells[row, 3].Value = item.납품장소;
                    sheet.Cells[row, 4].Value = item.품명;
                    sheet.Cells[row, 5].Value = item.규격;
                    sheet.Cells[row, 6].Value = item.세;
                    sheet.Cells[row, 7].Value = item.단위;
                    sheet.Cells[row, 8].Value = item.수량;
                    sheet.Cells[row, 9].Value = item.단가;
                    sheet.Cells[row, 10].Value = item.금액;
                    sheet.Cells[row, 11].Value = item.업체명;
                }

                // 열 너비 자동 조정
                sheet.Cells.AutoFitColumns();
            }

            private static void CreateSummarySheet(ExcelWorksheet sheet, List<TransData> data, Action<string>? log = null)
            {
                // 1. 식자재 납품 내역(계약) 테이블 생성
                sheet.Cells[1, 1].Value = "식자재 납품 내역(계약)";
                sheet.Cells[1, 1].Style.Font.Bold = true;
                sheet.Cells[1, 1].Style.Font.Size = 14;

                // 헤더
                sheet.Cells[2, 1].Value = "입고창고";
                sheet.Cells[2, 2].Value = "면세";
                sheet.Cells[2, 3].Value = "과세";
                sheet.Cells[2, 4].Value = "계";

                // 헤더 스타일
                using (var range = sheet.Cells[2, 1, 2, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                // 현장별 데이터 계산 (계약 데이터만)
                var contractData = data.Where(x => x.업체명 == "삼성웰스토리" && !x.납품장소.Contains("직원")).ToList();
                var warehouseGroups = contractData.GroupBy(x => x.납품장소).ToList();
                int row = 3;

                // 지정된 순서대로 현장 처리
                var orderedWarehouses = new[] { "양식뷔페", "양정식", "중식뷔페", "중정식", "한정식", "연회부", "운영지원부", "소모품" };

                foreach (var warehouse in orderedWarehouses)
                {
                    var group = warehouseGroups.FirstOrDefault(g => g.Key == warehouse);
                    var taxFreeAmount = group?.Where(x => x.부가세 == 0).Sum(x => x.금액) ?? 0;
                    var taxableAmount = group?.Where(x => x.부가세 > 0).Sum(x => x.금액 * 1.1) ?? 0; // 과세는 부가세 포함이므로 1.1 곱함
                    var totalAmount = taxFreeAmount + taxableAmount;

                    sheet.Cells[row, 1].Value = warehouse;
                    sheet.Cells[row, 2].Value = taxFreeAmount;
                    sheet.Cells[row, 3].Value = taxableAmount;
                    sheet.Cells[row, 4].Value = totalAmount;

                    row++;
                }

                // 2. 식자재 납품 내역(직원식당 계약) 테이블 생성
                row += 2; // 빈 행 추가
                sheet.Cells[row, 1].Value = "식자재 납품 내역(직원식당 계약)";
                sheet.Cells[row, 1].Style.Font.Bold = true;
                sheet.Cells[row, 1].Style.Font.Size = 14;

                row++;
                // 헤더
                sheet.Cells[row, 1].Value = "입고창고";
                sheet.Cells[row, 2].Value = "면세";
                sheet.Cells[row, 3].Value = "과세";
                sheet.Cells[row, 4].Value = "계";

                // 헤더 스타일
                using (var range = sheet.Cells[row, 1, row, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                row++;
                // 직원식당 데이터 (직원 데이터만)
                var employeeData = data.Where(x => x.업체명 == "삼성웰스토리" && x.납품장소.Contains("직원")).ToList();
                var employeeTaxFreeAmount = employeeData.Where(x => x.부가세 == 0).Sum(x => x.금액);
                var employeeTaxableAmount = employeeData.Where(x => x.부가세 > 0).Sum(x => x.금액 * 1.1); // 과세는 부가세 포함이므로 1.1 곱함
                var employeeTotalAmount = employeeTaxFreeAmount + employeeTaxableAmount;

                sheet.Cells[row, 1].Value = "직원식당";
                sheet.Cells[row, 2].Value = employeeTaxFreeAmount;
                sheet.Cells[row, 3].Value = employeeTaxableAmount;
                sheet.Cells[row, 4].Value = employeeTotalAmount;

                // 열 너비 자동 조정
                sheet.Cells.AutoFitColumns();
            }
        }

        public class RawData
        {
            public string 순번 { get; set; } = string.Empty;
            public string 입고일자 { get; set; } = string.Empty;
            public string 품번 { get; set; } = string.Empty;
            public string 품명 { get; set; } = string.Empty;
            public string 규격 { get; set; } = string.Empty;
            public string 수량 { get; set; } = string.Empty;
            public string 단위 { get; set; } = string.Empty;
            public string 단가 { get; set; } = string.Empty;
            public string 금액 { get; set; } = string.Empty;
            public string 부가세 { get; set; } = string.Empty;
            public string 합계금액 { get; set; } = string.Empty;
            public string 거래처명 { get; set; } = string.Empty;
            public string 적요 { get; set; } = string.Empty;
            public string 특이사항 { get; set; } = string.Empty;
            public string 현장명 { get; set; } = string.Empty;
            public string PJT코드 { get; set; } = string.Empty;
            public string PJT명 { get; set; } = string.Empty;
            public string 입고창고 { get; set; } = string.Empty;
        }

        public class TransData
        {
            public string 구매일자 { get; set; } = string.Empty;
            public string 납품장소 { get; set; } = string.Empty;
            public string 품명 { get; set; } = string.Empty;
            public string 규격 { get; set; } = string.Empty;
            public string 세 { get; set; } = string.Empty;
            public string 단위 { get; set; } = string.Empty;
            public double 수량 { get; set; }
            public double 단가 { get; set; }
            public double 금액 { get; set; }
            public double 부가세 { get; set; }
            public string 업체명 { get; set; } = string.Empty;
        }
    }
} 