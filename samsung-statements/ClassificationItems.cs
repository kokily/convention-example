using OfficeOpenXml;
using System.IO;

namespace SamsungStatements
{
    public static class ClassificationItems
    {
        public static void ClassifyItems(List<TransData> target, ExcelPackage workBook, string inputPath, LogManager? logManager = null)
        {
            logManager?.LogMessage($"DEBUG: ClassificationItems 시작 ===================> TransData 개수: {target.Count}");

            var prevContract = new List<TransData>();
            var prevNonContract = new List<TransData>();

            foreach (var data in target)
            {
                if (data.업체명 == "삼성웰스토리")
                {
                    prevContract.Add(data);
                }
                else if (data.업체명 == "삼성웰스토리(비)")
                {
                    prevNonContract.Add(data);
                }
            }

            logManager?.LogMessage($"DEBUG: 1차 분류 업체명별. 계약: {prevContract.Count}, 비계약: {prevNonContract.Count}");

            // 직원식당 분류
            var contract = new List<TransData>();
            var nonContract = new List<TransData>();
            var employee = new List<TransData>();
            var nonEmployee = new List<TransData>();

            foreach (var data in prevContract)
            {
                if (data.납품장소 != "직원식당")
                {
                    contract.Add(data);
                }
                else
                {
                    employee.Add(data);
                }
            }

            foreach (var data in prevNonContract)
            {
                if (data.납품장소.Replace(" ", "").Trim().Contains("직원식당"))
                {
                    nonEmployee.Add(data);
                }
                else
                {
                    nonContract.Add(data);
                }
            }

            logManager?.LogMessage($"DEBUG: 2차 분류 계약: {contract.Count}, 비계약: {nonContract.Count}, 직원: {employee.Count}, 직원 비계약: {nonEmployee.Count}");

            if (contract.Count > 0)
            {
                SaveToExcel(contract, workBook, "계약 식자재", logManager);
            }

            if (nonContract.Count > 0)
            {
                SaveToExcel(nonContract, workBook, "비계약 식자재", logManager);
            }

            if (employee.Count > 0)
            {
                SaveToExcel(employee, workBook, "계약 직원 식자재", logManager);
            }

            if (nonEmployee.Count > 0)
            {
                SaveToExcel(nonEmployee, workBook, "비계약 직원식당 식자재", logManager);
            }

            // --- 종합 시트 추가 ---
            AddSummarySheetV2(target, workBook, logManager);

            var outputFileName = "결산서.xlsx";
            var outputPath = Path.Combine(Path.GetDirectoryName(inputPath) ?? "", outputFileName);

            logManager?.LogMessage($"DEBUG: 엑셀 파일 저장 시도: {outputPath}");

            try
            {
                workBook.SaveAs(new FileInfo(outputPath));
                logManager?.LogMessage($"DEBUG: 엑셀파일 '{outputPath}' 저장 완료");
            }
            catch (Exception ex)
            {
                logManager?.LogMessage($"ERROR: 결과 엑셀 파일 저장 실패: {ex.Message}");
                throw new InvalidOperationException($"결과 엑셀 파일 저장 실패: {ex.Message}", ex);
            }
        }

        private static void SaveToExcel(List<TransData> dataList, ExcelPackage workBook, string sheetName, LogManager? logManager = null)
        {
            try
            {
                // 기존 시트가 있으면 삭제
                var existingWorksheet = workBook.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
                if (existingWorksheet != null)
                {
                    workBook.Workbook.Worksheets.Delete(existingWorksheet);
                }

                // 새 시트 생성
                var worksheet = workBook.Workbook.Worksheets.Add(sheetName);

                // 헤더 추가
                worksheet.Cells[1, 1].Value = "구매일자";
                worksheet.Cells[1, 2].Value = "납품장소";
                worksheet.Cells[1, 3].Value = "품명";
                worksheet.Cells[1, 4].Value = "규격";
                worksheet.Cells[1, 5].Value = "세";
                worksheet.Cells[1, 6].Value = "단위";
                worksheet.Cells[1, 7].Value = "수량";
                worksheet.Cells[1, 8].Value = "단가";
                worksheet.Cells[1, 9].Value = "금액";
                worksheet.Cells[1, 10].Value = "업체명";

                // 헤더 스타일 적용
                using (var range = worksheet.Cells[1, 1, 1, 10])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                // 데이터 추가
                for (int i = 0; i < dataList.Count; i++)
                {
                    var data = dataList[i];
                    int row = i + 2; // 헤더 다음 행부터

                    worksheet.Cells[row, 1].Value = data.구매일자;
                    worksheet.Cells[row, 2].Value = data.납품장소;
                    worksheet.Cells[row, 3].Value = data.품명;
                    worksheet.Cells[row, 4].Value = data.규격;
                    worksheet.Cells[row, 5].Value = data.세;
                    worksheet.Cells[row, 6].Value = data.단위;
                    worksheet.Cells[row, 7].Value = data.수량;
                    worksheet.Cells[row, 8].Value = data.단가;
                    worksheet.Cells[row, 9].Value = data.금액;
                    worksheet.Cells[row, 10].Value = data.업체명;
                }

                // 열 너비 자동 조정
                worksheet.Cells.AutoFitColumns();

                logManager?.LogMessage($"DEBUG: '{sheetName}' 시트에 {dataList.Count}개 데이터 저장 완료");
            }
            catch (Exception ex)
            {
                logManager?.LogMessage($"ERROR: '{sheetName}' 시트 저장 실패: {ex.Message}");
                throw new InvalidOperationException($"'{sheetName}' 시트 저장 실패: {ex.Message}", ex);
            }
        }

        private static void AddSummarySheetV2(List<TransData> allData, ExcelPackage workBook, LogManager? logManager = null)
        {
            var sheetName = "종합";
            var existingWorksheet = workBook.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
            if (existingWorksheet != null)
                workBook.Workbook.Worksheets.Delete(existingWorksheet);
            var ws = workBook.Workbook.Worksheets.Add(sheetName);

            // 표 행(카테고리) 고정
            string[] categories = new[] { "입고창고", "양식뷔페", "양정식", "중식뷔페", "중정식", "한정식", "연회부", "운영지원부", "소모품" };
            string[] empCategories = new[] { "입고창고", "직원식당" };
            string[] columns = new[] { "면세", "과세", "계" };

            // 표별 데이터 필터
            var contract = allData.Where(d => !d.납품장소.Trim().EndsWith("-비계약") && d.납품장소 != "직원식당").ToList();
            var empContract = allData.Where(d => d.납품장소 == "직원식당").ToList();

            int startRow = 1;
            // 1. 계약
            startRow = WriteSummaryTable(ws, startRow, "식자재 납품 내역(계약)", categories, contract);
            startRow += 2;
            // 2. 직원식당 계약
            startRow = WriteSummaryTable(ws, startRow, "식자재 납품 내역(직원식당 계약)", empCategories, empContract);

            ws.Cells.AutoFitColumns();
            logManager?.LogMessage($"DEBUG: '종합' 시트(이미지형) 추가 완료");
        }

        // 표 하나를 그리는 함수
        private static int WriteSummaryTable(ExcelWorksheet ws, int startRow, string title, string[] categories, List<TransData> data)
        {
            string[] columns = new[] { "면세", "과세", "계" };
            int nRows = categories.Length;
            int nCols = columns.Length;

            // 제목
            ws.Cells[startRow, 1].Value = title;
            ws.Cells[startRow, 1, startRow, nCols + 1].Merge = true;
            ws.Cells[startRow, 1, startRow, nCols + 1].Style.Font.Bold = true;
            ws.Cells[startRow, 1, startRow, nCols + 1].Style.Font.Size = 13;

            // 헤더 (A:입고창고, B:면세, C:과세, D:계)
            ws.Cells[startRow + 1, 1].Value = "입고창고";
            for (int i = 0; i < nCols; i++)
                ws.Cells[startRow + 1, i + 2].Value = columns[i];
            ws.Cells[startRow + 1, 1, startRow + 1, nCols + 1].Style.Font.Bold = true;

            // 데이터 행에서 '입고창고' 행 제거 (카테고리 첫 번째 행은 데이터로 사용하지 않음)
            int dataStartIdx = 1; // categories[1]부터 시작
            int dataRowCount = nRows - 1;

            double[] sumExempt = new double[dataRowCount];
            double[] sumTaxed = new double[dataRowCount];
            for (int i = 0; i < dataRowCount; i++)
            {
                string cat = categories[i + 1];
                IEnumerable<TransData> catData;
                if (cat == "소모품")
                {
                    catData = data.Where(d => d.납품장소.Replace("-비계약", "").StartsWith("소모품"));
                }
                else
                {
                    catData = data.Where(d => d.납품장소.Replace("-비계약", "") == cat);
                }
                sumExempt[i] = catData.Where(d => d.세 == "면").Sum(d => d.금액);
                sumTaxed[i] = catData.Where(d => d.세 == "과").Sum(d => d.금액);
            }

            // 표 채우기
            for (int i = 0; i < dataRowCount; i++)
            {
                ws.Cells[startRow + 2 + i, 1].Value = categories[i + 1];
                ws.Cells[startRow + 2 + i, 2].Value = sumExempt[i] == 0 ? 0 : sumExempt[i];
                ws.Cells[startRow + 2 + i, 3].Value = sumTaxed[i] == 0 ? 0 : sumTaxed[i];
                ws.Cells[startRow + 2 + i, 4].Value = (sumExempt[i] + sumTaxed[i]) == 0 ? 0 : (sumExempt[i] + sumTaxed[i]);
            }

            // 스타일: 데이터 행까지만 적용
            int lastDataRow = startRow + 2 + dataRowCount - 1;
            ws.Cells[startRow, 1, lastDataRow, 4].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            ws.Cells[startRow, 1, lastDataRow, 4].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            return lastDataRow + 1;
        }
    }
} 