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

                double 수량 = 0;
                if (!double.TryParse(item.수량, out 수량))
                {
                    logManager?.LogMessage($"WARN: 수량 파싱 오류 (행 {index + 1}): 값: '{item.수량}', 0으로 처리");
                    수량 = 0;
                }

                double 단가 = 0;
                if (!double.TryParse(item.단가, out 단가))
                {
                    logManager?.LogMessage($"WARN: 단가 파싱 오류 (행 {index + 1}): 값: '{item.단가}', 0으로 처리");
                    단가 = 0;
                }

                double 금액 = 0;
                if (!double.TryParse(item.금액, out 금액))
                {
                    logManager?.LogMessage($"WARN: 금액 파싱 오류 (행 {index + 1}): 값: '{item.금액}', 0으로 처리");
                    금액 = 0;
                }

                double 부가세 = 0;
                if (!double.TryParse(item.부가세, out 부가세))
                {
                    logManager?.LogMessage($"WARN: 부가세 파싱 오류 (행 {index + 1}): 값: '{item.부가세}', 0으로 처리");
                    부가세 = 0;
                }

                double 합계금액 = 0;
                if (!double.TryParse(item.합계금액, out 합계금액))
                {
                    logManager?.LogMessage($"WARN: 합계금액 파싱 오류 (행 {index + 1}): 값: '{item.합계금액}', 0으로 처리");
                    합계금액 = 0;
                }

                double 사용할금액 = 부가세 > 0 ? 합계금액 : 금액;

                // 디버그: 과세 품목의 금액 계산 확인 (처음 10개만)
                if (index < 10 && 부가세 > 0)
                {
                    logManager?.LogMessage($"DEBUG: 과세 품목 금액 계산 - 품명: {item.품명}, 일반금액: {금액:N0}, 부가세: {부가세:N0}, 합계금액: {합계금액:N0}, 사용할금액: {사용할금액:N0}");
                }

                double serialDate = 0;
                if (!double.TryParse(item.입고일자, out serialDate))
                {
                    logManager?.LogMessage($"WARN: 입고일자 파싱 오류 (행 {index + 1}): 값: '{item.입고일자}', 0으로 처리");
                    serialDate = 0;
                }

                var baseDate = new DateTime(1899, 12, 30);
                if (serialDate < 60)
                {
                    baseDate = new DateTime(1899, 12, 31);
                }
                var convertedTime = baseDate.AddDays(serialDate);
                var formattedDate = convertedTime.ToString("M월 d일", CultureInfo.GetCultureInfo("ko-KR"));

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
                    금액 = 사용할금액,
                    업체명 = item.거래처명
                };

                sources.Add(transItem);
                // 100개 단위로 진행률 표시
                if ((index + 1) % 100 == 0)
                {
                    logManager?.LogMessage($"DEBUG: 데이터 변환 진행률: {index + 1}/{target.Count} 건 처리 완료");
                }
            }

            logManager?.LogMessage($"DEBUG: ManufactureJson 종료. 변환된 데이터 수량 : {sources.Count}");
            logManager?.LogMessage($"DEBUG: ManufactureJson 종료 ===> 변환된 TransData 수: {sources.Count}");
            
            // 통계 정보 추가
            var uniqueSuppliers = sources.Select(s => s.업체명).Distinct().Count();
            var uniqueItems = sources.Select(s => s.품명).Distinct().Count();
            var totalAmount = sources.Sum(s => s.금액);
            var exemptAmount = sources.Where(s => s.세 == "면").Sum(s => s.금액);
            var taxedAmount = sources.Where(s => s.세 == "과").Sum(s => s.금액);
            var exemptCount = sources.Count(s => s.세 == "면");
            var taxedCount = sources.Count(s => s.세 == "과");
            
            logManager?.LogMessage($"DEBUG: 통계 - 업체 수: {uniqueSuppliers}, 품목 수: {uniqueItems}, 총 금액: {totalAmount:N0}원");
            logManager?.LogMessage($"DEBUG: 면세 통계 - 건수: {exemptCount}, 금액: {exemptAmount:N0}원");
            logManager?.LogMessage($"DEBUG: 과세 통계 - 건수: {taxedCount}, 금액: {taxedAmount:N0}원");
            
            return sources;
        }
    }
} 