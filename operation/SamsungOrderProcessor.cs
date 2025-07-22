using OfficeOpenXml;
using System.Text.RegularExpressions;
using System.Globalization;
using System.IO;

namespace IntegratedApp
{
    public static class SamsungOrderProcessor
    {
        public static void Process(string inputPath, string outputPath, Action<string>? log = null)
        {
            log?.Invoke("삼성웰스토리 발주서 전처리 시작...");
            
            // EPPlus 라이센스 설정
            ExcelPackage.License.SetNonCommercialPersonal("김현성");
            
            try
            {
                // 1단계: RawData 추출
                var transData = ExtractRawDataFromExcel(inputPath, log);
                log?.Invoke($"디버그: RawData 추출 및 변환 완료. 총 TransData 행 수: {transData.Count}");

                if (transData.Count == 0)
                {
                    throw new Exception("원본 파일에서 읽을 데이터가 없습니다.");
                }

                // 2단계: 데이터 변환
                var resultData = TransformData(transData, log);
                log?.Invoke($"디버그: 데이터 변환 완료. 총 ResultData 행 수: {resultData.Count}");

                // 3단계: 엑셀 파일 생성
                WriteProcessedDataToExcel(resultData, inputPath, outputPath, log);
                
                log?.Invoke("디버그: 엑셀 파일 저장 완료");
            }
            catch (Exception ex)
            {
                log?.Invoke($"오류 발생: {ex.Message}");
                throw;
            }
        }

        public static List<TransData> ExtractRawDataFromExcel(string filePath, Action<string>? log = null)
        {
            log?.Invoke($"디버그: 엑셀 변환 시작==============================\n파일: {filePath}");

            var transedData = new List<TransData>();

            using var package = new ExcelPackage(new FileInfo(filePath));
            
            if (package.Workbook.Worksheets.Count == 0)
            {
                throw new InvalidOperationException("엑셀 파일에 시트가 없습니다");
            }

            var worksheet = package.Workbook.Worksheets[0];
            var sheetName = worksheet.Name;

            // 병합된 셀 처리
            ProcessMergedCells(worksheet, log);

            // 실제 각 행 데이터 가져오기 -> 7행부터 실데이터
            int headerSkipRows = 6;
            int totalRows = worksheet.Dimension?.Rows ?? 0;

            for (int rIdx = 0; rIdx < totalRows; rIdx++)
            {
                if (rIdx < headerSkipRows)
                {
                    continue;
                }

                string aColVal = GetCellValue(worksheet, rIdx + 1, 1)?.Trim() ?? "";
                string bColVal = GetCellValue(worksheet, rIdx + 1, 2)?.Trim() ?? "";
                string cColVal = GetCellValue(worksheet, rIdx + 1, 3)?.Trim() ?? "";

                if (aColVal == "합계" || aColVal == "1/1" || cColVal == "소계")
                {
                    log?.Invoke($"디버그: 합계/소계/1/1 라인 스킵 (행 {rIdx + 1}): A:'{aColVal}' B:'{bColVal}' C:'{cColVal}'");
                    continue;
                }

                var data = new RawData
                {
                    SoldTo = aColVal,
                    사업장명 = GetCellValue(worksheet, rIdx + 1, 2) ?? "",
                    일자 = "",
                    품목코드 = GetCellValue(worksheet, rIdx + 1, 4) ?? "",
                    품목명 = GetCellValue(worksheet, rIdx + 1, 5) ?? "",
                    규격 = GetCellValue(worksheet, rIdx + 1, 6) ?? "",
                    단위 = GetCellValue(worksheet, rIdx + 1, 7) ?? "",
                    수량 = GetCellValue(worksheet, rIdx + 1, 8) ?? "",
                    단가 = GetCellValue(worksheet, rIdx + 1, 9) ?? "",
                    입고금액 = GetCellValue(worksheet, rIdx + 1, 10) ?? "",
                    부가세 = GetCellValue(worksheet, rIdx + 1, 11) ?? "",
                    합계 = GetCellValue(worksheet, rIdx + 1, 12) ?? ""
                };

                double parsed수량;
                try
                {
                    parsed수량 = Utils.ParseFloat(data.수량);
                }
                catch (Exception ex)
                {
                    throw new FormatException($"수량 파싱 오류: {ex.Message} (값: {data.수량})");
                }

                double parsed단가;
                try
                {
                    parsed단가 = Utils.ParseFloat(data.단가);
                }
                catch (Exception ex)
                {
                    throw new FormatException($"단가 파싱 오류: {ex.Message} (값: {data.단가})");
                }

                double parsed입고금액;
                try
                {
                    parsed입고금액 = Utils.ParseFloat(data.입고금액);
                }
                catch (Exception ex)
                {
                    throw new FormatException($"입고금액 파싱 오류: {ex.Message} (값: {data.입고금액})");
                }

                double parsed부가세;
                try
                {
                    parsed부가세 = Utils.ParseFloat(data.부가세);
                }
                catch (Exception ex)
                {
                    throw new FormatException($"부가세 파싱 오류: {ex.Message} (값: {data.부가세})");
                }

                var trans = new TransData
                {
                    사업장명 = data.사업장명,
                    품목코드 = data.품목코드,
                    품목명 = data.품목명,
                    규격 = data.규격,
                    바코드 = "",
                    수량 = parsed수량,
                    평균단가 = parsed단가,
                    입고금액 = parsed입고금액,
                    부가세 = parsed부가세,
                    단위 = data.단위
                };

                transedData.Add(trans);
            }

            log?.Invoke($"디버그: RawData 추출 완료. 행: {transedData.Count}");
            return transedData;
        }

        private static void ProcessMergedCells(ExcelWorksheet worksheet, Action<string>? log = null)
        {
            try
            {
                foreach (var mergedCell in worksheet.MergedCells)
                {
                    var match = Regex.Match(mergedCell, @"^([A-Z]+)(\d+):([A-Z]+)(\d+)$");
                    if (!match.Success) continue;

                    string startCol = match.Groups[1].Value;
                    int startRow = int.Parse(match.Groups[2].Value);
                    string endCol = match.Groups[3].Value;
                    int endRow = int.Parse(match.Groups[4].Value);

                    if (startCol.ToUpper() != "B") continue;

                    string startCell = $"{startCol}{startRow}";
                    string val = GetCellValue(worksheet, startRow, GetColumnIndex(startCol))?.Trim() ?? "";

                    if (string.IsNullOrEmpty(val)) continue;

                    // 병합된 셀의 값을 모든 행에 전파
                    for (int r = startRow; r <= endRow; r++)
                    {
                        worksheet.Cells[r, 2].Value = val;
                    }

                    log?.Invoke($"디버그: 병합 셀 '{startCell}' 값 전파 완료: '{val}' (행 {startRow}~{endRow})");
                }
            }
            catch (Exception ex)
            {
                log?.Invoke($"WARN: 병합된 셀 정보 처리 실패: {ex.Message}");
            }
        }

        private static string? GetCellValue(ExcelWorksheet worksheet, int row, int col)
        {
            try
            {
                var cell = worksheet.Cells[row, col];
                return cell.Value?.ToString();
            }
            catch
            {
                return null;
            }
        }

        private static int GetColumnIndex(string columnName)
        {
            int index = 0;
            foreach (char c in columnName.ToUpper())
            {
                index = index * 26 + (c - 'A' + 1);
            }
            return index;
        }

        public static List<TransData> TransformData(List<TransData> transData, Action<string>? log = null)
        {
            log?.Invoke($"디버그: 트랜스 데이터 시작=====================> 데이터 행: {transData.Count}");

            var resultData = new List<TransData>();

            foreach (var item in transData)
            {
                // 사업장명 가공
                string parsed사업장명 = item.사업장명.TrimStart("국방컨벤션(".ToCharArray());
                parsed사업장명 = parsed사업장명.TrimEnd(')');
                parsed사업장명 = parsed사업장명.Replace("/", "-");
                parsed사업장명 = parsed사업장명 + "-" + Utils.Iif(item.부가세 == 0, "면", "과");

                var processedRow = new TransData
                {
                    사업장명 = parsed사업장명,
                    품목코드 = "25" + item.품목코드,
                    품목명 = item.품목명,
                    규격 = item.규격,
                    바코드 = "",
                    수량 = item.수량,
                    평균단가 = item.평균단가,
                    입고금액 = item.입고금액,
                    부가세 = item.부가세,
                    단위 = item.단위
                };

                resultData.Add(processedRow);
            }

            log?.Invoke($"디버깅: 트랜스폼 완료. 총 {resultData.Count} 행");
            return resultData;
        }

        public static void WriteProcessedDataToExcel(List<TransData> data, string inputFilePath, string outputPath, Action<string>? log = null)
        {
            log?.Invoke($"디버그: 엑셀파일 저장 시작==================> 데이터 행 수: {data.Count}");

            using var package = new ExcelPackage();

            // 총괄시트 생성
            string sheetNameTotal = "총괄";
            var totalSheet = package.Workbook.Worksheets.Add(sheetNameTotal);

            if (data.Count > 0)
            {
                AddSingleSheetDataOnly(totalSheet, data);
            }
            else
            {
                log?.Invoke("WARN: 처리할 엑셀 데이터가 없습니다.");
            }

            // 데이터 분류 시작
            var classifiedData = data.GroupBy(item => item.사업장명)
                                   .ToDictionary(group => group.Key, group => group.ToList());

            // 지정된 사업장 순서
            string[] orderedBusinessUnits = {
                "양식뷔페",
                "양식뷔페-비계약",
                "양식소모품",
                "양식소모품-비계약",
                "양정식",
                "양정식-비계약",
                "연회부",
                "연회부-비계약",
                "연회부소모품",
                "연회부소모품-비계약",
                "운영지원부",
                "운영지원부-비계약",
                "중식뷔페",
                "중식뷔페-비계약",
                "중식소모품",
                "중식소모품-비계약",
                "중정식",
                "중정식-비계약",
                "직원식당",
                "직원식당-비계약",
                "한정식",
                "한정식-비계약"
            };

            // 지정된 순서대로 사업장별 시트 생성 및 데이터 추가
            foreach (string businessUnit in orderedBusinessUnits)
            {
                if (classifiedData.ContainsKey(businessUnit))
                {
                    var sheetData = classifiedData[businessUnit];
                    log?.Invoke($"디버그: 시트 '{businessUnit}' 생성 및 데이터 추가 시작, 행 수: {sheetData.Count}");
                    var worksheet = package.Workbook.Worksheets.Add(businessUnit);
                    AddSingleSheetDataOnly(worksheet, sheetData);
                }
            }
            
            // orderedBusinessUnits에 없는 사업장명도 추가
            foreach (var key in classifiedData.Keys)
            {
                if (!orderedBusinessUnits.Contains(key))
                {
                    var sheetData = classifiedData[key];
                    log?.Invoke($"디버그: 시트 '{key}'(추가) 생성 및 데이터 추가 시작, 행 수: {sheetData.Count}");
                    var worksheet = package.Workbook.Worksheets.Add(key);
                    AddSingleSheetDataOnly(worksheet, sheetData);
                }
            }

            log?.Invoke($"디버그: 엑셀파일 저장 시도: {outputPath}");

            try
            {
                package.SaveAs(new FileInfo(outputPath));
                log?.Invoke($"디버그: 엑셀파일 '{outputPath}' 저장 완료");
            }
            catch (Exception ex)
            {
                throw new Exception($"결과 엑셀파일 저장 실패: {ex.Message}", ex);
            }
        }

        private static void AddSingleSheetDataOnly(ExcelWorksheet worksheet, List<TransData> data)
        {
            string[] headers = { "사업장명", "품목코드", "품목명", "규격", "바코드", "수량", "평균단가", "입고금액", "부가세", "", "단위" };

            // 헤더 추가
            for (int colIdx = 0; colIdx < headers.Length; colIdx++)
            {
                worksheet.Cells[1, colIdx + 1].Value = headers[colIdx];
            }

            // 데이터 추가
            for (int rowIdx = 0; rowIdx < data.Count; rowIdx++)
            {
                var item = data[rowIdx];
                int row = rowIdx + 2;

                worksheet.Cells[row, 1].Value = item.사업장명;
                worksheet.Cells[row, 2].Value = item.품목코드;
                worksheet.Cells[row, 3].Value = item.품목명;
                worksheet.Cells[row, 4].Value = item.규격;
                worksheet.Cells[row, 5].Value = item.바코드;
                worksheet.Cells[row, 6].Value = item.수량;
                worksheet.Cells[row, 7].Value = item.평균단가;
                worksheet.Cells[row, 8].Value = item.입고금액;
                worksheet.Cells[row, 9].Value = item.부가세;
                // J열(10번째)은 비워둠
                worksheet.Cells[row, 11].Value = item.단위;
            }
        }

        public class RawData
        {
            public string SoldTo { get; set; } = string.Empty;
            public string 사업장명 { get; set; } = string.Empty;
            public string 일자 { get; set; } = string.Empty;
            public string 품목코드 { get; set; } = string.Empty;
            public string 품목명 { get; set; } = string.Empty;
            public string 규격 { get; set; } = string.Empty;
            public string 단위 { get; set; } = string.Empty;
            public string 수량 { get; set; } = string.Empty;
            public string 단가 { get; set; } = string.Empty;
            public string 입고금액 { get; set; } = string.Empty;
            public string 부가세 { get; set; } = string.Empty;
            public string 합계 { get; set; } = string.Empty;
        }

        public class TransData
        {
            public string 사업장명 { get; set; } = string.Empty;
            public string 품목코드 { get; set; } = string.Empty;
            public string 품목명 { get; set; } = string.Empty;
            public string 규격 { get; set; } = string.Empty;
            public string 바코드 { get; set; } = string.Empty;
            public double 수량 { get; set; }
            public double 평균단가 { get; set; }
            public double 입고금액 { get; set; }
            public double 부가세 { get; set; }
            public string 단위 { get; set; } = string.Empty;
        }

        public static class Utils
        {
            public static double ParseFloat(string s)
            {
                if (string.IsNullOrEmpty(s))
                {
                    return 0;
                }

                s = s.Replace(",", "");

                if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double val))
                {
                    return val;
                }

                throw new FormatException($"'{s}'를 double로 변환 실패");
            }

            public static string Iif(bool condition, string trueVal, string falseVal)
            {
                return condition ? trueVal : falseVal;
            }
        }
    }
} 