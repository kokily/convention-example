using System;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Text;
using System.IO;

namespace EtcOrder
{
    public static class CsvToJson
    {
        public static List<RawData> ConvertCsvToJson(string filePath)
        {
            var jsonData = new List<RawData>();

            try
            {
                ExcelPackage.License.SetNonCommercialPersonal("김현성");

                using var package = new ExcelPackage(new FileInfo(filePath));
                var worksheet = package.Workbook.Worksheets[0];

                if (worksheet == null)
                    throw new InvalidOperationException("엑셀 파일에 시트가 없습니다");

                var dataRange = worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column, 
                                               worksheet.Dimension.End.Row, worksheet.Dimension.End.Column];
                var rows = dataRange.Value as object[,];
                if (rows == null)
                    throw new InvalidOperationException("엑셀 파일에서 데이터를 읽을 수 없습니다.");

                int rowCount = rows.GetLength(0);
                if (rowCount < 2)
                    throw new InvalidOperationException("엑셀 파일에 데이터가 충분하지 않습니다.");

                for (int rIdx = 0; rIdx < rowCount; rIdx++)
                {
                    var aColVal = GetCellValue(rows, rIdx, 0)?.ToString()?.Trim() ?? "";
                    if (aColVal == "기간별구매현황[일자별]" || aColVal == "순번")
                        continue;

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
                }
                return jsonData;
            }
            catch (Exception ex)
            {
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