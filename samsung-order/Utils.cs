using System;
using System.Globalization;
using System.Collections.Generic;

namespace SamsungOrder
{
    public static class Utils
    {
        public static double ParseFloat(string s)
        {
            if (string.IsNullOrEmpty(s))
            {
                return 0;
            }

            s = s.Replace(",", "");

            if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double val))
            {
                return val;
            }

            throw new FormatException($"'{s}'를 double로 변환 실패");
        }

        public static string Iif(bool condition, string trueVal, string falseVal)
        {
            return condition ? trueVal : falseVal;
        }

        public static string GetCellValue(string[] row, int index)
        {
            if (index >= 0 && index < row.Length)
            {
                return row[index];
            }

            return "";
        }

        public static List<TransData> TransformData(List<TransData> transData)
        {
            Console.WriteLine($"디버그: 트랜스 데이터 시작=====================> 데이터 행: {transData.Count}");

            var resultData = new List<TransData>();

            foreach (var item in transData)
            {
                // 사업장명 가공
                string parsed사업장명 = item.사업장명.TrimStart("국방컨벤션(".ToCharArray());
                parsed사업장명 = parsed사업장명.TrimEnd(')');
                parsed사업장명 = parsed사업장명.Replace("/", "-");
                parsed사업장명 = parsed사업장명 + "-" + Iif(item.부가세 == 0, "면", "과");

                var processedRow = new TransData
                {
                    사업장명 = parsed사업장명,
                    품목코드 = "25" + item.품목코드,
                    품목명 = item.품목명,
                    규격 = item.규격,
                    바코드 = "",
                    수량 = item.수량,
                    평균단가 = item.평균단가,
                    입고금액 = item.입고금액,
                    부가세 = item.부가세,
                    단위 = item.단위
                };

                resultData.Add(processedRow);
            }

            Console.WriteLine($"디버깅: 트랜스폼 완료. 총 {resultData.Count} 행");

            return resultData;
        }
    }
} 