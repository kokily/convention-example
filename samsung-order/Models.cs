namespace SamsungOrder
{
    public class RawData
    {
        public string SoldTo { get; set; } = string.Empty;
        public string 사업장명 { get; set; } = string.Empty;
        public string 일자 { get; set; } = string.Empty;
        public string 품목코드 { get; set; } = string.Empty;
        public string 품목명 { get; set; } = string.Empty;
        public string 규격 { get; set; } = string.Empty;
        public string 단위 { get; set; } = string.Empty;
        public string 수량 { get; set; } = string.Empty;
        public string 단가 { get; set; } = string.Empty;
        public string 입고금액 { get; set; } = string.Empty;
        public string 부가세 { get; set; } = string.Empty;
        public string 합계 { get; set; } = string.Empty;
    }

    public class TransData
    {
        public string 사업장명 { get; set; } = string.Empty;
        public string 품목코드 { get; set; } = string.Empty;
        public string 품목명 { get; set; } = string.Empty;
        public string 규격 { get; set; } = string.Empty;
        public string 바코드 { get; set; } = string.Empty;
        public double 수량 { get; set; }
        public double 평균단가 { get; set; }
        public double 입고금액 { get; set; }
        public double 부가세 { get; set; }
        public string 단위 { get; set; } = string.Empty;
    }
} 