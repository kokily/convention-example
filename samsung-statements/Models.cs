namespace SamsungStatements
{
    public class RawData
    {
        public string 순번 { get; set; } = string.Empty;
        public string 입고일자 { get; set; } = string.Empty;
        public string 품번 { get; set; } = string.Empty;
        public string 품명 { get; set; } = string.Empty;
        public string 규격 { get; set; } = string.Empty;
        public string 수량 { get; set; } = string.Empty;
        public string 단위 { get; set; } = string.Empty;
        public string 단가 { get; set; } = string.Empty;
        public string 금액 { get; set; } = string.Empty;
        public string 부가세 { get; set; } = string.Empty;
        public string 합계금액 { get; set; } = string.Empty;
        public string 거래처명 { get; set; } = string.Empty;
        public string 적요 { get; set; } = string.Empty;
        public string 특이사항 { get; set; } = string.Empty;
        public string 현장명 { get; set; } = string.Empty;
        public string PJT코드 { get; set; } = string.Empty;
        public string PJT명 { get; set; } = string.Empty;
        public string 입고창고 { get; set; } = string.Empty;
    }

    public class TransData
    {
        public string 구매일자 { get; set; } = string.Empty;
        public string 납품장소 { get; set; } = string.Empty;
        public string 품명 { get; set; } = string.Empty;
        public string 규격 { get; set; } = string.Empty;
        public string 세 { get; set; } = string.Empty;
        public string 단위 { get; set; } = string.Empty;
        public double 수량 { get; set; }
        public double 단가 { get; set; }
        public double 금액 { get; set; }
        public string 업체명 { get; set; } = string.Empty;
    }
} 