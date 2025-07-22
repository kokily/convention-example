using System;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IntegratedApp
{
    public static class ExpendablesProcessor
    {
        public static void Process(string inputPath, string outputPath, Action<string>? log = null)
        {
            log?.Invoke("컨벤션 소모품 결산 처리 시작...");
            
            ExcelPackage.License.SetNonCommercialPersonal("김현성");
            
            try
            {
                var sourceList = new List<SourceItem>();
                using (var package = new ExcelPackage(new FileInfo(inputPath)))
                {
                    var ws = package.Workbook.Worksheets[0];
                    int rowCount = ws.Dimension.End.Row;
                    int colCount = ws.Dimension.End.Column;
                    var headerMap = new Dictionary<string, int>();
                    int 추정금액_재고단가_Col = -1;
                    
                    for (int c = 1; c <= colCount; c++)
                    {
                        var h2 = ws.Cells[2, c].Text.Trim();
                        var h3 = ws.Cells[3, c].Text.Trim();
                        if (h2 == "①추정금액" && h3 == "재고단가")
                            추정금액_재고단가_Col = c;
                        if (!string.IsNullOrEmpty(h2) && !string.IsNullOrEmpty(h3))
                            headerMap[h2 + "_" + h3] = c;
                        if (!string.IsNullOrEmpty(h2))
                            headerMap[h2] = c;
                        if (!string.IsNullOrEmpty(h3))
                            headerMap[h3] = c;
                    }
                    
                    if (추정금액_재고단가_Col == -1)
                        throw new Exception("엑셀에서 ①추정금액/재고단가 열을 찾을 수 없습니다.");
                    
                    for (int r = 4; r <= rowCount; r++)
                    {
                        if (!headerMap.ContainsKey("창고"))
                        {
                            log?.Invoke($"[ERROR] '창고' 키가 headerMap에 없습니다. 실제 키 목록: {string.Join(", ", headerMap.Keys)}");
                            throw new Exception("headerMap에 '창고' 키가 없습니다.");
                        }
                        
                        string 창고명 = ws.Cells[r, headerMap["창고"]].Text;
                        // 소모삭제(전체)는 아예 제외
                        if (창고명.Contains("소모삭제(전체)")) continue;
                        if (창고명.Contains("합계")) continue;
                        
                        var item = new SourceItem
                        {
                            순번 = ws.Cells[r, headerMap["순번"]].Text,
                            창고 = 창고명,
                            계정구분 = ws.Cells[r, headerMap["계정구분"]].Text,
                            품번 = ws.Cells[r, headerMap["품번"]].Text,
                            품명 = ws.Cells[r, headerMap["품명"]].Text,
                            규격 = ws.Cells[r, headerMap["규격"]].Text,
                            단위 = ws.Cells[r, headerMap["단위"]].Text,
                            기초수량 = ws.Cells[r, headerMap["기초수량"]].Text,
                            입고수량 = ws.Cells[r, headerMap["입고수량"]].Text,
                            출고수량 = ws.Cells[r, headerMap["출고수량"]].Text,
                            재고수량 = ws.Cells[r, headerMap["재고수량"]].Text,
                            재고단가 = ws.Cells[r, 추정금액_재고단가_Col].Text,
                            주거래처 = ws.Cells[r, headerMap["주거래처"]].Text,
                            바코드 = ws.Cells[r, headerMap["바코드"]].Text
                        };
                        sourceList.Add(item);
                    }
                    log?.Invoke($"총 {sourceList.Count}건 데이터 파싱 완료");
                }
                
                // 창고명 매핑 및 (비) 창고명 처리
                foreach (var item in sourceList)
                {
                    string originalWarehouse = item.창고;
                    item.창고 = MapWarehouse(item.창고, item.품번);
                    if (item.품번 == "2321011001" && item.품명 == "건해삼")
                    {
                        log?.Invoke($"[DEBUG] 품목: {item.품번} {item.품명} - 원본창고: {originalWarehouse} -> 매핑창고: {item.창고}");
                    }
                }
                log?.Invoke("창고명 매핑 완료");
                
                // 창고 순서(비 포함)
                var warehouseOrder = new[] {
                    "운영지원부", "운영지원부(비)", "예약팀", "연회실", "연회실(비)", "연회주류", "연회주류(비)", "양식뷔페", "양식뷔페(비)", "양정식", "양정식(비)", "양식별도",
                    "중식뷔페", "중식뷔페(비)", "중정식", "중정식(비)", "중식별도", "한정식", "한정식(비)", "직원식당", "직원식당(비)", "소모품(조리실)", "소모품(조리실)(비)", "시설관리팀"
                };
                
                // 창고별 그룹화
                var grouped = sourceList
                    .Where(x => warehouseOrder.Contains(x.창고))
                    .GroupBy(x => x.창고)
                    .OrderBy(g => Array.IndexOf(warehouseOrder, g.Key));
                
                log?.Invoke($"창고별 그룹화 완료: {grouped.Count()}개 창고");
                
                using (var outPkg = new ExcelPackage())
                {
                    // 1. 총괄 시트 생성
                    var allRows = new List<object[]>();
                    string[] headers = { "순번", "창고", "품번", "품명", "규격", "단위", "이월수량", "입고수량", "출고수량", "재고수량", "단가", "재고금액" };
                    allRows.Add(headers);
                    int totalRowNum = 1;
                    
                    foreach (var group in grouped)
                    {
                        foreach (var item in group)
                        {
                            totalRowNum++;
                            var resultItem = new ResultItem
                            {
                                순번 = totalRowNum - 1,
                                창고 = item.창고,
                                품번 = item.품번,
                                품명 = item.품명,
                                규격 = item.규격,
                                단위 = item.단위,
                                이월수량 = Utils.ParseDouble(item.기초수량),
                                입고수량 = Utils.ParseDouble(item.입고수량),
                                출고수량 = Utils.ParseDouble(item.출고수량),
                                재고수량 = Utils.ParseDouble(item.재고수량),
                                단가 = Utils.ParseDouble(item.재고단가),
                                재고금액 = Utils.ParseDouble(item.재고수량) * Utils.ParseDouble(item.재고단가)
                            };
                            
                            allRows.Add(new object[] {
                                resultItem.순번, resultItem.창고, resultItem.품번, resultItem.품명, resultItem.규격, resultItem.단위,
                                resultItem.이월수량, resultItem.입고수량, resultItem.출고수량, resultItem.재고수량, resultItem.단가, resultItem.재고금액
                            });
                        }
                    }
                    
                    var totalSheet = outPkg.Workbook.Worksheets.Add("총괄");
                    totalSheet.Cells[1, 1].LoadFromArrays(allRows.ToArray());
                    
                    // 헤더 스타일
                    using (var range = totalSheet.Cells[1, 1, 1, 12])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }
                    
                    // 열 너비 자동 조정
                    totalSheet.Cells.AutoFitColumns();
                    
                    // 2. 창고별 시트 생성
                    foreach (var group in grouped)
                    {
                        var sheetName = group.Key;
                        foreach (var c in Path.GetInvalidFileNameChars())
                            sheetName = sheetName.Replace(c, '_');
                        
                        if (string.IsNullOrWhiteSpace(sheetName)) continue;
                        
                        var sheet = outPkg.Workbook.Worksheets.Add(sheetName);
                        
                        // 헤더 설정
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
                        
                        int row = 2;
                        double totalCarryOver = 0, totalIn = 0, totalOut = 0, totalStock = 0, totalAmount = 0;
                        
                        foreach (var item in group)
                        {
                            var resultItem = new ResultItem
                            {
                                순번 = row - 1,
                                창고 = item.창고,
                                품번 = item.품번,
                                품명 = item.품명,
                                규격 = item.규격,
                                단위 = item.단위,
                                이월수량 = Utils.ParseDouble(item.기초수량),
                                입고수량 = Utils.ParseDouble(item.입고수량),
                                출고수량 = Utils.ParseDouble(item.출고수량),
                                재고수량 = Utils.ParseDouble(item.재고수량),
                                단가 = Utils.ParseDouble(item.재고단가),
                                재고금액 = Utils.ParseDouble(item.재고수량) * Utils.ParseDouble(item.재고단가)
                            };
                            
                            sheet.Cells[row, 1].Value = resultItem.순번;
                            sheet.Cells[row, 2].Value = resultItem.창고;
                            sheet.Cells[row, 3].Value = resultItem.품번;
                            sheet.Cells[row, 4].Value = resultItem.품명;
                            sheet.Cells[row, 5].Value = resultItem.규격;
                            sheet.Cells[row, 6].Value = resultItem.단위;
                            sheet.Cells[row, 7].Value = resultItem.이월수량;
                            sheet.Cells[row, 8].Value = resultItem.입고수량;
                            sheet.Cells[row, 9].Value = resultItem.출고수량;
                            sheet.Cells[row, 10].Value = resultItem.재고수량;
                            sheet.Cells[row, 11].Value = resultItem.단가;
                            sheet.Cells[row, 12].Value = resultItem.재고금액;
                            
                            totalCarryOver += resultItem.이월수량;
                            totalIn += resultItem.입고수량;
                            totalOut += resultItem.출고수량;
                            totalStock += resultItem.재고수량;
                            totalAmount += resultItem.재고금액;
                            
                            row++;
                        }
                        
                        // 합계 행 추가
                        sheet.Cells[row, 1].Value = "합계";
                        sheet.Cells[row, 7].Value = totalCarryOver;
                        sheet.Cells[row, 8].Value = totalIn;
                        sheet.Cells[row, 9].Value = totalOut;
                        sheet.Cells[row, 10].Value = totalStock;
                        sheet.Cells[row, 12].Value = totalAmount;
                        
                        // 합계 행 스타일
                        using (var range = sheet.Cells[row, 1, row, headers.Length])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
                        }
                        
                        // 열 너비 자동 조정
                        sheet.Cells.AutoFitColumns();
                        
                        log?.Invoke($"시트 '{group.Key}' 생성 완료: {group.Count()}건");
                    }
                    
                    outPkg.SaveAs(new FileInfo(outputPath));
                    log?.Invoke($"컨벤션 소모품 결산 완료: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                log?.Invoke($"오류 발생: {ex.Message}");
                throw;
            }
        }

        private static string MapWarehouse(string src, string 품번)
        {
            // 할인 창고 처리
            if (IsDiscountWarehouseOrigin(src))
            {
                double discountRate = GetDiscountRate(품번);
                if (discountRate > 0)
                {
                    return src + "(비)";
                }
            }
            return src;
        }

        private static bool IsDiscountWarehouseOrigin(string warehouse)
        {
            var discountWarehouses = new[] { "운영지원부", "연회실", "연회주류", "양식뷔페", "양정식", "중식뷔페", "중정식", "한정식", "직원식당", "소모품(조리실)" };
            return discountWarehouses.Contains(warehouse);
        }

        private static double GetDiscountRate(string 품번)
        {
            // 할인율 매핑 (실제 비즈니스 로직에 맞게 수정 필요)
            var discountItems = new[] { "2321011001", "2321011002", "2321011003" }; // 예시 품번
            return discountItems.Contains(품번) ? 0.1 : 0; // 10% 할인
        }

        public static class Utils
        {
            public static double ParseDouble(string value)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return 0;

                // 쉼표 제거
                value = value.Replace(",", "");
                
                if (double.TryParse(value, out double result))
                    return result;
                
                return 0;
            }
        }

        public class SourceItem
        {
            public string 순번 { get; set; } = string.Empty;
            public string 창고 { get; set; } = string.Empty;
            public string 계정구분 { get; set; } = string.Empty;
            public string 품번 { get; set; } = string.Empty;
            public string 품명 { get; set; } = string.Empty;
            public string 규격 { get; set; } = string.Empty;
            public string 단위 { get; set; } = string.Empty;
            public string 기초수량 { get; set; } = string.Empty;
            public string 입고수량 { get; set; } = string.Empty;
            public string 출고수량 { get; set; } = string.Empty;
            public string 재고수량 { get; set; } = string.Empty;
            public string 재고단가 { get; set; } = string.Empty;
            public string 주거래처 { get; set; } = string.Empty;
            public string 바코드 { get; set; } = string.Empty;
        }

        public class ResultItem
        {
            public int 순번 { get; set; }
            public string 창고 { get; set; } = string.Empty;
            public string 품번 { get; set; } = string.Empty;
            public string 품명 { get; set; } = string.Empty;
            public string 규격 { get; set; } = string.Empty;
            public string 단위 { get; set; } = string.Empty;
            public double 이월수량 { get; set; }
            public double 입고수량 { get; set; }
            public double 출고수량 { get; set; }
            public double 재고수량 { get; set; }
            public double 단가 { get; set; }
            public double 재고금액 { get; set; }
        }
    }
} 