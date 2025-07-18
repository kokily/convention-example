namespace expendables_excel_converter
{
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
        public string 재고단가 { get; set; } = string.Empty; // ①추정금액_재고단가
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