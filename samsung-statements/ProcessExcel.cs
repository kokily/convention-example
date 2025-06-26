using OfficeOpenXml;
using System.IO;

namespace SamsungStatements
{
    public static class ProcessExcel
    {
        public static void ProcessExcelFile(string inputPath, LogManager? logManager = null)
        {
            try
            {
                // 1단계: CsvToJson - 엑셀 파일을 RawData로 변환
                var rawData = CsvToJson.ConvertCsvToJson(inputPath, logManager);

                logManager?.LogMessage($"DEBUG: CsvToJson 완료 ===> 로드 RawData 행: {rawData.Count}");

                if (rawData.Count == 0)
                {
                    throw new InvalidOperationException("원본 엑셀파일의 데이터가 부족합니다.");
                }

                // 2단계: ManufactureJson - RawData를 TransData로 변환
                var inputFileName = Path.GetFileNameWithoutExtension(inputPath);
                var transData = ManufactureJson.ManufactureJsonData(rawData, inputFileName, logManager);

                logManager?.LogMessage($"DEBUG: ManufactureJson 완료 ===> 변환된 TransData 수: {transData.Count}");

                if (transData.Count == 0)
                {
                    throw new InvalidOperationException("데이터 변환 후 유효한 데이터가 부족합니다.");
                }

                // 3단계: 새로운 Excel 워크북 생성
                // EPPlus 라이센스 설정
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using var workBook = new ExcelPackage();

                // 기본 Sheet1 삭제 시도
                try
                {
                    var defaultSheet = workBook.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Sheet1");
                    if (defaultSheet != null)
                    {
                        workBook.Workbook.Worksheets.Delete(defaultSheet);
                        logManager?.LogMessage("DEBUG: 기본 Sheet1 삭제 완료");
                    }
                }
                catch (Exception ex)
                {
                    logManager?.LogMessage($"WARN: Sheet1 삭제 실패 (정상동작일 수 있음): {ex.Message}");
                }

                // 4단계: 데이터 분류 및 시트 추가
                ClassificationItems.ClassifyItems(transData, workBook, inputPath, logManager);

                logManager?.LogMessage("DEBUG: 분류 시트들 추가 및 Excel 파일 저장 완료");
            }
            catch (Exception ex)
            {
                logManager?.LogMessage($"ERROR: ProcessExcelFile 오류: {ex.Message}");
                throw new InvalidOperationException($"ProcessExcelFile 오류: {ex.Message}", ex);
            }
        }
    }
} 