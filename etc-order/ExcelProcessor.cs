using System;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;

namespace EtcOrder
{
    public static class ExcelProcessor
    {
        public static void Process(string inputPath, string outputPath, System.Action<string>? log = null)
        {
            var rawList = CsvToJson.ConvertCsvToJson(inputPath);
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
                        else if (DateTime.TryParseExact(inDate, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var dt2))
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
                            var roundedUnitPrice = Math.Round(v2);
                            sheet.Cells[rowIdx, 8].Value = roundedUnitPrice;
                        }
                        else
                        {
                            sheet.Cells[rowIdx, 8].Value = item.단가;
                        }
                        
                        // 금액 처리 (정수로 반올림)
                        if (double.TryParse(item.금액.Replace(",", ""), out var v3))
                        {
                            var roundedAmount = Math.Round(v3);
                            sheet.Cells[rowIdx, 9].Value = roundedAmount;
                            totalAmount += roundedAmount;
                        }
                        else
                        {
                            sheet.Cells[rowIdx, 9].Value = item.금액;
                        }
                        
                        // 부가세 처리 (정수로 반올림)
                        if (double.TryParse(item.부가세.Replace(",", ""), out var v4))
                        {
                            var roundedVAT = Math.Round(v4);
                            sheet.Cells[rowIdx, 10].Value = roundedVAT;
                            totalVAT += roundedVAT;
                        }
                        else
                        {
                            sheet.Cells[rowIdx, 10].Value = item.부가세;
                        }
                        
                        // 합계금액 처리 (정수로 반올림)
                        if (double.TryParse(item.합계금액.Replace(",", ""), out var v5))
                        {
                            var roundedSum = Math.Round(v5);
                            sheet.Cells[rowIdx, 11].Value = roundedSum;
                            totalSum += roundedSum;
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

                    // "계" 행 추가 (마지막 데이터 행보다 2-3행 아래)
                    int summaryRow = rowIdx + 2;
                    sheet.Cells[summaryRow, 5].Value = "계";
                    sheet.Cells[summaryRow, 6].Value = Math.Round(totalQuantity);
                    sheet.Cells[summaryRow, 9].Value = Math.Round(totalAmount);
                    sheet.Cells[summaryRow, 10].Value = Math.Round(totalVAT);
                    sheet.Cells[summaryRow, 11].Value = Math.Round(totalSum);

                    // 서식 지정
                    foreach (int col in new[] {1, 3, 4, 5, 7, 12, 13, 14})
                        sheet.Column(col).Style.Numberformat.Format = "@";
                    foreach (int col in new[] {6, 8, 9, 10, 11})
                        sheet.Column(col).Style.Numberformat.Format = "#,##0";
                    sheet.Column(2).Style.Numberformat.Format = "@";

                    sheet.Cells.AutoFitColumns();
                    log?.Invoke($"[{sheetName}] 처리 완료");
                }
                outPkg.SaveAs(new FileInfo(outputPath));
                log?.Invoke($"엑셀 저장 완료: {outputPath}");
            }
        }
    }
} 