using SupportTuoiBeo.Data;

namespace SupportTuoiBeo.Models
{
    public class UserDetailViewModel
    {
        public int Id { get; set; }

        public string MaKH { get; set; }

        public string Tinh { get; set; }

        public long TienThanhToan { get; set; }

        public int Ngay { get; set; }

        public EnumThang Thang { get; set; }
    }
}