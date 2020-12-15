namespace SupportTuoiBeo.Models
{
    public class TotalViewModel
    {
        public int Id { get; set; }

        public string MaKH { get; set; }

        public string Tinh { get; set; }

        public long DoanhThuThang1 { get; set; }

        public long DoanhThuThang2 { get; set; }

        public long DoanhThuThang3 { get; set; }

        public long DoanhThuThang4 { get; set; }

        public long DoanhThuThang5 { get; set; }

        public long DoanhThuThang6 { get; set; }

        public long DoanhThuThang7 { get; set; }

        public long DoanhThuThang8 { get; set; }

        public long DoanhThuThang9 { get; set; }

        public long DoanhThuThang10 { get; set; }

        public long DoanhThuThang11 { get; set; }

        public long DoanhThuThang12 { get; set; }

        public long Total
        {
            get
            {
                return
  DoanhThuThang1 + DoanhThuThang2 + DoanhThuThang3 + DoanhThuThang4 + DoanhThuThang5 + DoanhThuThang6
  + DoanhThuThang7 + DoanhThuThang8 + DoanhThuThang9 + DoanhThuThang10 + DoanhThuThang11 + DoanhThuThang12;
            }
        }
    }
}