using System.Globalization;

namespace SamsungStatements
{
    public static class ManufactureJson
    {
        public static List<TransData> ManufactureJsonData(List<RawData> target, string inputFileName, LogManager? logManager = null)
        {
            var sources = new List<TransData>();

            for (int index = 0; index < target.Count; index++)
            {
                var item = target[index];

                if (!double.TryParse(item.수량, out double 수량))
                {
                    logManager?.LogMessage($"ERROR: ManufactureJson - 수량 파싱 오류 (행 {index + 1}): 값: '{item.수량}'");
                    throw new InvalidOperationException($"수량 파싱 오류: 값: {item.수량}");
                }

                if (!double.TryParse(item.단가, out double 단가))
                {
                    logManager?.LogMessage($"ERROR: ManufactureJson - 단가 파싱 오류 (행 {index + 1}): 값: '{item.단가}'");
                    throw new InvalidOperationException($"단가 파싱 오류: 값: {item.단가}");
                }

                if (!double.TryParse(item.금액, out double 금액))
                {
                    logManager?.LogMessage($"ERROR: ManufactureJson - 금액 파싱 오류 (행 {index + 1}): 값: '{item.금액}'");
                    throw new InvalidOperationException($"금액 파싱 오류: 값: {item.금액}");
                }

                if (!double.TryParse(item.부가세, out double 부가세))
                {
                    logManager?.LogMessage($"ERROR: ManufactureJson - 부가세 파싱 오류 (행 {index + 1}): 값: '{item.부가세}'");
                    throw new InvalidOperationException($"부가세 파싱 오류: 값: {item.부가세}");
                }

                // 구매일자 변환 로직
                if (!double.TryParse(item.입고일자, out double serialDate))
                {
                    logManager?.LogMessage($"ERROR: ManufactureJson - 입고일자 파싱 오류 (행 {index + 1}): 값: '{item.입고일자}'");
                    throw new InvalidOperationException($"입고일자 파싱 오류: 값: {item.입고일자}");
                }

                var baseDate = new DateTime(1899, 12, 30);

                if (serialDate < 60)
                {
                    baseDate = new DateTime(1899, 12, 31);
                }

                var convertedTime = baseDate.AddDays(serialDate);
                var formattedDate = convertedTime.ToString("M월 d일", CultureInfo.GetCultureInfo("ko-KR"));

                // 납품장소 변환 로직
                var processedPlace = item.현장명;

                if (processedPlace.Contains("-비계약"))
                {
                    processedPlace = processedPlace.Replace("-비계약", "");
                    logManager?.LogMessage($"DEBUG: -비계약 삭제 {processedPlace}");
                }

                switch (processedPlace)
                {
                    case "소모품(양식당)":
                        processedPlace = "소모품(양식)";
                        logManager?.LogMessage($"DEBUG: 납품장소 '{item.현장명}' -> '{processedPlace}' (by Pattern)");
                        break;
                    case "소모품(중식당)":
                        processedPlace = "소모품(중식)";
                        logManager?.LogMessage($"DEBUG: 납품장소 '{item.현장명}' -> '{processedPlace}' (by Pattern)");
                        break;
                    case "소모품(연회부)":
                        processedPlace = "소모품(연회)";
                        logManager?.LogMessage($"DEBUG: 납품장소 '{item.현장명}' -> '{processedPlace}' (by Pattern)");
                        break;
                    case "소모품(양식당-비)":
                        processedPlace = "소모품(양식)";
                        logManager?.LogMessage($"DEBUG: 납품장소 '{item.현장명}' -> '{processedPlace}' (by Pattern)");
                        break;
                    case "소모품(중식당-비)":
                        processedPlace = "소모품(중식)";
                        logManager?.LogMessage($"DEBUG: 납품장소 '{item.현장명}' -> '{processedPlace}' (by Pattern)");
                        break;
                    case "소모품(연회부-비)":
                        processedPlace = "소모품(연회)";
                        logManager?.LogMessage($"DEBUG: 납품장소 '{item.현장명}' -> '{processedPlace}' (by Pattern)");
                        break;
                }

                var transItem = new TransData
                {
                    구매일자 = formattedDate,
                    납품장소 = processedPlace,
                    품명 = item.품명,
                    규격 = item.규격,
                    세 = 부가세 == 0 ? "면" : "과",
                    단위 = item.단위,
                    수량 = 수량,
                    단가 = 단가,
                    금액 = 금액,
                    업체명 = item.거래처명
                };

                sources.Add(transItem);
                logManager?.LogMessage($"DEBUG: ManufactureJson {index + 1}번째 데이터 추가: 품명='{transItem.품명}'");
            }

            return sources;
        }
    }
} 