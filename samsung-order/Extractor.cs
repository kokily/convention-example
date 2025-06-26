using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace SamsungOrder
{
    public static class Extractor
    {
        public static List<TransData> ExtractRawDataFromExcel(string filePath)
        {
            Console.WriteLine($"디버그: 엑셀 변환 시작==============================\n파일: {filePath}");

            var transedData = new List<TransData>();

            // EPPlus 라이센스 설정
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(filePath));
            
            if (package.Workbook.Worksheets.Count == 0)
            {
                throw new InvalidOperationException("엑셀 파일에 시트가 없습니다");
            }

            var worksheet = package.Workbook.Worksheets[0];
            var sheetName = worksheet.Name;

            // 병합된 셀 처리
            ProcessMergedCells(worksheet);

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
                    Console.WriteLine($"디버그: 합계/소계/1/1 라인 스킵 (행 {rIdx + 1}): A:'{aColVal}' B:'{bColVal}' C:'{cColVal}'");
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

            Console.WriteLine($"디버그: RawData 추출 완료. 행: {transedData.Count}");

            return transedData;
        }

        private static void ProcessMergedCells(ExcelWorksheet worksheet)
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

                    Console.WriteLine($"디버그: 병합 셀 '{startCell}' 값 전파 완료: '{val}' (행 {startRow}~{endRow})");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"WARN: 병합된 셀 정보 처리 실패: {ex.Message}");
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
    }
} 