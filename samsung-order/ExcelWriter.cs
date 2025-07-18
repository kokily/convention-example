using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SamsungOrder
{
    public static class ExcelWriter
    {
        public static void WriteProcessedDataToExcel(List<TransData> data, string inputFilePath)
        {
            Console.WriteLine($"디버그: 엑셀파일 저장 시작==================> 데이터 행 수: {data.Count}");

            // EPPlus 라이센스 설정
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

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
                Console.WriteLine("WARN: 처리할 엑셀 데이터가 없습니다.");
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
                    Console.WriteLine($"디버그: 시트 '{businessUnit}' 생성 및 데이터 추가 시작, 행 수: {sheetData.Count}");
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
                    Console.WriteLine($"디버그: 시트 '{key}'(추가) 생성 및 데이터 추가 시작, 행 수: {sheetData.Count}");
                    var worksheet = package.Workbook.Worksheets.Add(key);
                    AddSingleSheetDataOnly(worksheet, sheetData);
                }
            }

            // 결과물 엑셀 파일 저장
            string fileName = Path.GetFileNameWithoutExtension(inputFilePath);
            string outputFileName = $"분류된_{fileName}.xlsx";
            string outputPath = Path.Combine(Path.GetDirectoryName(inputFilePath)!, outputFileName);

            Console.WriteLine($"디버그: 엑셀파일 저장 시도: {outputPath}");

            try
            {
                package.SaveAs(new FileInfo(outputPath));
                Console.WriteLine($"디버그: 엑셀파일 '{outputPath}' 저장 완료");
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
    }
} 