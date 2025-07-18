using System;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace expendables_excel_converter
{
    public static class ExcelProcessor
    {
        public static void Process(string inputPath, string outputPath, Action<string>? log = null)
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
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
                        log?.Invoke($"[{group.Key}] 시트 데이터 {group.Count()}건");
                        var sorted = group.OrderBy(x => x.품명).ToList();
                        foreach (var item in sorted)
                        {
                            double.TryParse(item.기초수량.Replace(",", ""), out double 이월수량);
                            double.TryParse(item.입고수량.Replace(",", ""), out double 입고수량);
                            double.TryParse(item.출고수량.Replace(",", ""), out double 출고수량);
                            double.TryParse(item.재고수량.Replace(",", ""), out double 재고수량);
                            double.TryParse(item.재고단가.Replace(",", ""), out double 단가);
                            double 재고금액 = 0;
                            if (IsDiscountWarehouseOrigin(item.창고))
                            {
                                double rate = GetDiscountRate(item.품번);
                                재고금액 = Math.Round(단가 * rate * 재고수량, 0, MidpointRounding.AwayFromZero);
                            }
                            else
                            {
                                재고금액 = Math.Round(단가 * 재고수량, 0, MidpointRounding.AwayFromZero);
                            }
                            재고금액 = RoundByFive(재고금액);
                            object[] rowData = {
                                totalRowNum, item.창고, item.품번, item.품명, item.규격, item.단위,
                                이월수량, 입고수량, 출고수량, 재고수량, 단가, 재고금액
                            };
                            allRows.Add(rowData);
                            totalRowNum++;
                        }
                    }
                    var wsTotal = outPkg.Workbook.Worksheets.Add("총괄");
                    for (int i = 0; i < allRows.Count; i++)
                        for (int j = 0; j < headers.Length; j++)
                            wsTotal.Cells[i + 1, j + 1].Value = allRows[i][j];
                    wsTotal.Cells[1, 1, allRows.Count, headers.Length].AutoFitColumns();
                    log?.Invoke("총괄 시트 생성 완료");
                    // 2. 창고별 시트 생성
                    foreach (var group in grouped)
                    {
                        if (group.Count() == 0) continue;
                        var ws = outPkg.Workbook.Worksheets.Add(group.Key);
                        for (int i = 0; i < headers.Length; i++)
                            ws.Cells[1, i + 1].Value = headers[i];
                        var sorted = group.OrderBy(x => x.품명).ToList();
                        int row = 2;
                        foreach (var item in sorted)
                        {
                            double.TryParse(item.기초수량.Replace(",", ""), out double 이월수량);
                            double.TryParse(item.입고수량.Replace(",", ""), out double 입고수량);
                            double.TryParse(item.출고수량.Replace(",", ""), out double 출고수량);
                            double.TryParse(item.재고수량.Replace(",", ""), out double 재고수량);
                            double.TryParse(item.재고단가.Replace(",", ""), out double 단가);
                            double 재고금액 = 0;
                            if (IsDiscountWarehouseOrigin(item.창고))
                            {
                                double rate = GetDiscountRate(item.품번);
                                재고금액 = Math.Round(단가 * rate * 재고수량, 0, MidpointRounding.AwayFromZero);
                            }
                            else
                            {
                                재고금액 = Math.Round(단가 * 재고수량, 0, MidpointRounding.AwayFromZero);
                            }
                            재고금액 = RoundByFive(재고금액);
                            ws.Cells[row, 1].Value = row - 1;
                            ws.Cells[row, 2].Value = item.창고;
                            ws.Cells[row, 3].Value = item.품번;
                            ws.Cells[row, 4].Value = item.품명;
                            ws.Cells[row, 5].Value = item.규격;
                            ws.Cells[row, 6].Value = item.단위;
                            ws.Cells[row, 7].Value = 이월수량;
                            ws.Cells[row, 8].Value = 입고수량;
                            ws.Cells[row, 9].Value = 출고수량;
                            ws.Cells[row, 10].Value = 재고수량;
                            ws.Cells[row, 11].Value = 단가;
                            ws.Cells[row, 12].Value = 재고금액;
                            row++;
                        }
                        ws.Cells[1, 1, row - 1, headers.Length].AutoFitColumns();
                        log?.Invoke($"[{group.Key}] 시트 생성 완료");
                    }
                    log?.Invoke($"엑셀 저장 시도: {outputPath}");
                    outPkg.SaveAs(new FileInfo(outputPath));
                    log?.Invoke($"엑셀 저장 완료: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                log?.Invoke($"[ERROR] 변환 중 예외 발생: {ex.Message}\n{ex.StackTrace}");
            }
        }
        // 창고명 매핑 함수, 낙찰률 창고명 구분 함수 등은 아래에 이어서 구현
        // 창고명 매핑 규칙 (질문 5항 전체 반영)
        private static string MapWarehouse(string src, string 품번)
        {
            if (string.IsNullOrWhiteSpace(src)) return src;
            
            // (비) 창고명 매핑 - 먼저 처리
            if (src.Contains("식재료(운영부-비)")) return "운영지원부(비)";
            if (src.Contains("식재료(연회부-비)")) return "연회실(비)";
            if (src.Contains("소모품(연회부-비)")) return "연회실(비)";
            if (src.Contains("식재료(연회주류-비)")) return "연회주류(비)";
            if (src.Contains("식재료(양식뷔페-비)")) return "양식뷔페(비)";
            if (src.Contains("식재료(중식뷔페-비)")) return "중식뷔페(비)";
            if (src.Contains("식재료(양정식-비)")) return "양정식(비)";
            if (src.Contains("식재료(중정식-비)")) return "중정식(비)";
            if (src.Contains("식재료(한정식-비)")) return "한정식(비)";
            if (src.Contains("식재료(직원식당-비)")) return "직원식당(비)";
            if (src.Contains("소모품(양식당-비)")) return "소모품(조리실)(비)";
            if (src.Contains("소모품(중식당-비)")) return "소모품(조리실)(비)";
            
            // 일반 창고명 매핑
            if (src.Contains("국방컨벤션") || src.Contains("국방컨벤션(재료)") || src.Contains("피복비(운영부)")) return "운영지원부";
            if (src.Contains("수용비(예약실)") || src.Contains("식재료비(예약실)") || src.Contains("일반재료비(예약실)")) return "예약팀";
            if (src.Contains("수용비(연회부)") || src.Contains("피복비(연회부)") || src.Contains("식재료비(연회부)") || src.Contains("일반재료비(연회부)") || src.Contains("소모품(연회부)")) return "연회실";
            if (src.Contains("식재료비(주류)")) return "연회주류";
            if (src.Contains("수용비(양식뷔페)") || src.Contains("식재료비(양식뷔페)") || src.Contains("일반재료비(양식뷔페)")) return "양식뷔페";
            if (src.Contains("수용비(중식뷔페)") || src.Contains("식재료비(중식뷔페)")) return "중식뷔페";
            if (src.Contains("식재료비(양정식)")) return "양정식";
            if (src.Contains("식재료비(중정식)")) return "중정식";
            if (src.Contains("식재료비(한식당)") || src.Contains("식재료비(한정식)")) return "한정식";
            if (src.Contains("식재료비(직원)")) return "직원식당";
            if (src.Contains("수용비(한식당)") || src.Contains("피복비(양식당)") || src.Contains("소모품(양식당)") || src.Contains("소모품(중식당)")) return "소모품(조리실)";
            if (src.Contains("양식당(기타)")) return "양식별도";
            if (src.Contains("중식당(기타)")) return "중식별도";
            if (src.Contains("수용비(시설팀)") || src.Contains("피복비(시설팀)") || src.Contains("시설장비유지(전기)") || src.Contains("시설유지(영선/기계)") || src.Contains("시설유지(기계)") || src.Contains("시설유지(소방)") || src.Contains("시설유지(공구)")) return "시설관리팀";
            return src;
        }
        // 낙찰률 적용 대상 창고명 판별
        private static bool IsDiscountWarehouseOrigin(string warehouse)
        {
            string[] targets = {
                "양식뷔페(비)", "양정식(비)", "중식뷔페(비)", "중정식(비)", "한정식(비)", "직원식당(비)", "연회실(비)", "연회주류(비)", "운영지원부(비)", "소모품(조리실)(비)"
            };
            return targets.Any(t => warehouse.Contains(t));
        }
        // 품번 앞 2자리로 낙찰률 반환
        private static double GetDiscountRate(string 품번)
        {
            if (string.IsNullOrEmpty(품번) || 품번.Length < 2) return 1.0;
            string prefix = 품번.Substring(0, 2);
            return prefix switch
            {
                "20" => 0.789,
                "21" => 0.734,
                "22" => 0.92324,
                "23" => 0.88405,
                "24" => 0.6997,
                "25" => 0.71042,
                _ => 1.0
            };
        }
        // 5 기준 절상/절사
        private static double RoundByFive(double value)
        {
            int intVal = (int)value;
            int mod = intVal % 10;
            if (mod == 0) return intVal;
            if (mod < 5) return intVal - mod;
            return intVal - mod + 10;
        }
    }
} 