using System;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;

namespace IntegratedApp
{
    public static class EtcOrderProcessor
    {
        public static void Process(string inputPath, string outputPath, Action<string>? log = null)
        {
            log?.Invoke("거래처별 엑셀 변환기 시작...");
            
            // EPPlus 라이센스 설정
            ExcelPackage.License.SetNonCommercialPersonal("김현성");
            
            var rawList = CsvToJson.ConvertCsvToJson(inputPath, log);
            
            // 1. 지정 업체명 리스트
            var customOrder = new[] {
                "더브레드뱅크", "에스엠발아커피", "미래자판기", "롯데칠성", "평생유통", "금호주류상사(유)", "광림통상", "디케이유통", "푸드가온 주식회사", "디엔케이드림", "더와이컴퍼니", "해피바이어스", "복지단용산마트", "한수상사"
            };
            
            var filtered = rawList.Where(x => x.거래처명 != "삼성웰스토리" && x.거래처명 != "삼성웰스토리(비)").ToList();
            var allNames = filtered.Where(x => !string.IsNullOrWhiteSpace(x.거래처명)).Select(x => x.거래처명).Distinct().ToList();
            
            // customOrder에 없는 나머지 업체명 가나다순
            var restNames = allNames.Except(customOrder).OrderBy(x => x).ToList();
            var finalOrder = customOrder.Concat(restNames).ToList();

            using (var outPkg = new ExcelPackage())
            {
                foreach (var sheetNameOrigin in finalOrder)
                {
                    var sheetName = sheetNameOrigin;
                    foreach (var c in Path.GetInvalidFileNameChars())
                        sheetName = sheetName.Replace(c, '_');
                    if (string.IsNullOrWhiteSpace(sheetName)) continue;

                    var group = filtered.Where(x => x.거래처명 == sheetNameOrigin).ToList();
                    int rowCount = group.Count();
                    double sumAmount = group.Sum(x => double.TryParse(x.합계금액.Replace(",", ""), out var v) ? v : 0);
                    log?.Invoke($"[{sheetName}] 처리 시작: {rowCount}건, 합계금액 {sumAmount:N0}");

                    var sheet = outPkg.Workbook.Worksheets.Add(sheetName);
                    
                    // 헤더 (불필요한 열 제거)
                    sheet.Cells[1, 1].Value = "순번";
                    sheet.Cells[1, 2].Value = "입고일자";
                    sheet.Cells[1, 3].Value = "품번";
                    sheet.Cells[1, 4].Value = "품명";
                    sheet.Cells[1, 5].Value = "규격";
                    sheet.Cells[1, 6].Value = "수량";
                    sheet.Cells[1, 7].Value = "단위";
                    sheet.Cells[1, 8].Value = "단가";
                    sheet.Cells[1, 9].Value = "금액";
                    sheet.Cells[1, 10].Value = "부가세";
                    sheet.Cells[1, 11].Value = "합계금액";
                    sheet.Cells[1, 12].Value = "거래처명";
                    sheet.Cells[1, 13].Value = "현장명";
                    sheet.Cells[1, 14].Value = "입고창고";

                    int rowIdx = 2;
                    double totalQuantity = 0;
                    double totalAmount = 0;
                    double totalVAT = 0;
                    double totalSum = 0;

                    foreach (var item in group)
                    {
                        // 입고일자 변환: yyyy-MM-dd, 시리얼 → M월 d일
                        string inDate = item.입고일자;
                        string dateStr = inDate;
                        if (DateTime.TryParse(inDate, out var dt))
                            dateStr = $"{dt.Month}월 {dt.Day}일";
                        else if (DateTime.TryParseExact(inDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt2))
                            dateStr = $"{dt2.Month}월 {dt2.Day}일";
                        else if (double.TryParse(inDate, out var serial))
                        {
                            var baseDate = new DateTime(1899, 12, 30);
                            var excelDate = baseDate.AddDays(serial);
                            dateStr = $"{excelDate.Month}월 {excelDate.Day}일";
                        }

                        sheet.Cells[rowIdx, 1].Value = item.순번;
                        sheet.Cells[rowIdx, 2].Value = dateStr;
                        sheet.Cells[rowIdx, 3].Value = item.품번;
                        sheet.Cells[rowIdx, 4].Value = item.품명;
                        sheet.Cells[rowIdx, 5].Value = item.규격;
                        
                        // 수량 처리 (정수로 반올림)
                        if (double.TryParse(item.수량.Replace(",", ""), out var v1))
                        {
                            var roundedQuantity = Math.Round(v1);
                            sheet.Cells[rowIdx, 6].Value = roundedQuantity;
                            totalQuantity += roundedQuantity;
                        }
                        else
                        {
                            sheet.Cells[rowIdx, 6].Value = item.수량;
                        }
                        
                        sheet.Cells[rowIdx, 7].Value = item.단위;
                        
                        // 단가 처리 (정수로 반올림)
                        if (double.TryParse(item.단가.Replace(",", ""), out var v2))
                        {
                            var roundedPrice = Math.Round(v2);
                            sheet.Cells[rowIdx, 8].Value = roundedPrice;
                        }
                        else
                        {
                            sheet.Cells[rowIdx, 8].Value = item.단가;
                        }
                        
                        // 금액 처리
                        if (double.TryParse(item.금액.Replace(",", ""), out var v3))
                        {
                            sheet.Cells[rowIdx, 9].Value = v3;
                            totalAmount += v3;
                        }
                        else
                        {
                            sheet.Cells[rowIdx, 9].Value = item.금액;
                        }
                        
                        // 부가세 처리
                        if (double.TryParse(item.부가세.Replace(",", ""), out var v4))
                        {
                            sheet.Cells[rowIdx, 10].Value = v4;
                            totalVAT += v4;
                        }
                        else
                        {
                            sheet.Cells[rowIdx, 10].Value = item.부가세;
                        }
                        
                        // 합계금액 처리
                        if (double.TryParse(item.합계금액.Replace(",", ""), out var v5))
                        {
                            sheet.Cells[rowIdx, 11].Value = v5;
                            totalSum += v5;
                        }
                        else
                        {
                            sheet.Cells[rowIdx, 11].Value = item.합계금액;
                        }
                        
                        sheet.Cells[rowIdx, 12].Value = item.거래처명;
                        sheet.Cells[rowIdx, 13].Value = item.현장명;
                        sheet.Cells[rowIdx, 14].Value = item.입고창고;
                        
                        rowIdx++;
                    }

                    // 합계 행 추가
                    sheet.Cells[rowIdx, 1].Value = "합계";
                    sheet.Cells[rowIdx, 6].Value = totalQuantity;
                    sheet.Cells[rowIdx, 9].Value = totalAmount;
                    sheet.Cells[rowIdx, 10].Value = totalVAT;
                    sheet.Cells[rowIdx, 11].Value = totalSum;

                    // 합계 행 스타일
                    using (var range = sheet.Cells[rowIdx, 1, rowIdx, 14])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
                    }

                    // 헤더 스타일
                    using (var range = sheet.Cells[1, 1, 1, 14])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }

                    // 열 너비 자동 조정
                    sheet.Cells.AutoFitColumns();

                    log?.Invoke($"[{sheetName}] 처리 완료: {rowCount}건");
                }

                outPkg.SaveAs(new FileInfo(outputPath));
                log?.Invoke($"거래처별 엑셀 변환 완료: {outputPath}");
            }
        }

        public static class CsvToJson
        {
            public static List<RawData> ConvertCsvToJson(string filePath, Action<string>? log = null)
            {
                log?.Invoke($"CSV 파일 읽기 시작: {filePath}");
                
                var rawDataList = new List<RawData>();
                
                using var package = new ExcelPackage(new FileInfo(filePath));
                var worksheet = package.Workbook.Worksheets[0];
                
                int totalRows = worksheet.Dimension?.Rows ?? 0;
                log?.Invoke($"총 {totalRows}행의 데이터를 읽습니다.");

                // 헤더는 건너뛰고 데이터만 읽기
                for (int row = 2; row <= totalRows; row++)
                {
                    var rawData = new RawData
                    {
                        순번 = GetCellValue(worksheet, row, 1),
                        입고일자 = GetCellValue(worksheet, row, 2),
                        품번 = GetCellValue(worksheet, row, 3),
                        품명 = GetCellValue(worksheet, row, 4),
                        규격 = GetCellValue(worksheet, row, 5),
                        수량 = GetCellValue(worksheet, row, 6),
                        단위 = GetCellValue(worksheet, row, 7),
                        단가 = GetCellValue(worksheet, row, 8),
                        금액 = GetCellValue(worksheet, row, 9),
                        부가세 = GetCellValue(worksheet, row, 10),
                        합계금액 = GetCellValue(worksheet, row, 11),
                        거래처명 = GetCellValue(worksheet, row, 12),
                        적요 = GetCellValue(worksheet, row, 13),
                        특이사항 = GetCellValue(worksheet, row, 14),
                        현장명 = GetCellValue(worksheet, row, 15),
                        PJT코드 = GetCellValue(worksheet, row, 16),
                        PJT명 = GetCellValue(worksheet, row, 17),
                        입고창고 = GetCellValue(worksheet, row, 18)
                    };

                    // 빈 행은 건너뛰기
                    if (!string.IsNullOrWhiteSpace(rawData.품명))
                    {
                        rawDataList.Add(rawData);
                    }
                }

                log?.Invoke($"CSV 변환 완료: {rawDataList.Count}건의 데이터를 읽었습니다.");
                return rawDataList;
            }

            private static string GetCellValue(ExcelWorksheet worksheet, int row, int col)
            {
                try
                {
                    var cell = worksheet.Cells[row, col];
                    return cell.Value?.ToString()?.Trim() ?? "";
                }
                catch
                {
                    return "";
                }
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
            public string 업체명 { get; set; } = string.Empty;
        }
    }
} 