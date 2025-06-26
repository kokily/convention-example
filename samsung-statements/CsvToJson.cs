using OfficeOpenXml;
using System.Text;
using System.IO;

namespace SamsungStatements
{
    public static class CsvToJson
    {
        public static List<RawData> ConvertCsvToJson(string filePath, LogManager? logManager = null)
        {
            logManager?.LogMessage($"디버그: CsvToJson 시작 =====================> 파일: {filePath}");

            var jsonData = new List<RawData>();

            try
            {
                // EPPlus 라이센스 설정
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using var package = new ExcelPackage(new FileInfo(filePath));
                var worksheet = package.Workbook.Worksheets[0]; // 첫 번째 시트

                if (worksheet == null)
                {
                    logManager?.LogMessage("ERROR: 엑셀 파일에 시트가 없습니다");
                    throw new InvalidOperationException("엑셀 파일에 시트가 없습니다");
                }

                logManager?.LogMessage($"DEBUG: 시트 이름 {worksheet.Name}");

                // 전체 데이터 범위 가져오기
                var dataRange = worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column, 
                                               worksheet.Dimension.End.Row, worksheet.Dimension.End.Column];
                
                var rows = dataRange.Value as object[,];
                
                if (rows == null)
                {
                    logManager?.LogMessage("ERROR: 엑셀 파일에서 데이터를 읽을 수 없습니다.");
                    throw new InvalidOperationException("엑셀 파일에서 데이터를 읽을 수 없습니다.");
                }

                int rowCount = rows.GetLength(0);
                int colCount = rows.GetLength(1);

                logManager?.LogMessage($"DEBUG: 읽은 총 행 수 (헤더포함) {rowCount}");

                if (rowCount < 2)
                {
                    logManager?.LogMessage("ERROR: 엑셀 파일에 데이터가 충분하지 않습니다.");
                    throw new InvalidOperationException("엑셀 파일에 데이터가 충분하지 않습니다.");
                }

                for (int rIdx = 0; rIdx < rowCount; rIdx++)
                {
                    var aColVal = GetCellValue(rows, rIdx, 0)?.ToString()?.Trim() ?? "";

                    if (aColVal == "기간별구매현황[일자별]" || aColVal == "순번")
                    {
                        logManager?.LogMessage($"DEBUG: 제목 라인 스킵 (행 {rIdx + 1})");
                        continue;
                    }

                    //                    var dColVal = GetCellValue(rows, rIdx, 3)?.ToString()?.Trim() ?? "";
                    //
                    //                    if (dColVal == "[ 합 계 ]")
                    //                    {
                    //                        logManager?.LogMessage($"DEBUG: [ 합 계 ] 라인 스킵 (행 {rIdx + 1})");
                    //                        continue;
                    //                    }

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
                        logManager?.LogMessage($"DEBUG: 엑셀 데이터 읽기 진행률: {rIdx + 1}/{rowCount} 행 처리 완료");
                    }
                }

                logManager?.LogMessage($"DEBUG: CsvToJson 종료. 변환된 데이터 수량 : {jsonData.Count}");

                return jsonData;
            }
            catch (Exception ex)
            {
                logManager?.LogMessage($"ERROR: 엑셀 파일 읽기 실패: {ex.Message}");
                throw new InvalidOperationException($"엑셀 파일 읽기 실패: {ex.Message}", ex);
            }
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
} 