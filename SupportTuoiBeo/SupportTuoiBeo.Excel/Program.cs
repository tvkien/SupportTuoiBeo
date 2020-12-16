using EnumsNET;
using SupportTuoiBeo.Data;
using SupportTuoiBeo.Data.Contexts;
using SupportTuoiBeo.Data.Entities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SupportTuoiBeo.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start export data");

            const string path = @"D:\QLBH-2020-Total.xlsx";
            var totals = GetTotalViewModels();
            ExportHelper.Export(path, totals);

            Console.WriteLine("Completed export data");

            Console.ReadKey();
        }

        private void RetrieveAndInsertData()
        {
            const string fileName = @"D:\QLBH-2020-Kien-Update-Convert.xlsx";

            foreach (Data.EnumThang thang in (Data.EnumThang[])Enum.GetValues(typeof(Data.EnumThang)))
            {
                if (thang == Data.EnumThang.All)
                {
                    continue;
                }
                string description = thang.AsString(EnumFormat.Description);

                var userDetails = ExportExtension.GetUserDetails(fileName, description, thang);

                InsertDatabase(userDetails);
            }
        }

        private static void InsertDatabase(List<UserDetails> userDetails)
        {
            using var db = new ApplicationDbContext();
            db.UserDetails.AddRange(userDetails);
            db.SaveChanges();
        }

        private static List<TotalViewModel> GetTotalViewModels()
        {
            var totals = new List<TotalViewModel>();
            var listUserDetail = new List<UserDetails>();

            using (var db = new ApplicationDbContext())
            {
                listUserDetail = db.UserDetails.ToList();
            }

            var listMaKH = listUserDetail.GroupBy(x => x.MaKH).Select(x => x.Key);
            var index = 1;

            foreach (var maKh in listMaKH)
            {
                var totalView = new TotalViewModel
                {
                    Id = index,
                    MaKH = maKh,
                    Tinh = listUserDetail.FirstOrDefault(x => x.MaKH == maKh).Tinh
                };
                var listOneKH = listUserDetail.Where(x => x.MaKH == maKh).ToList();
                foreach (var one in listOneKH)
                {
                    switch (one.Thang)
                    {
                        case EnumThang.Thang01:
                            totalView.DoanhThuThang1 += one.TienThanhToan;
                            break;
                        case EnumThang.Thang02:
                            totalView.DoanhThuThang2 += one.TienThanhToan;
                            break;
                        case EnumThang.Thang03:
                            totalView.DoanhThuThang3 += one.TienThanhToan;
                            break;
                        case EnumThang.Thang04:
                            totalView.DoanhThuThang4 += one.TienThanhToan;
                            break;
                        case EnumThang.Thang05:
                            totalView.DoanhThuThang5 += one.TienThanhToan;
                            break;
                        case EnumThang.Thang06:
                            totalView.DoanhThuThang6 += one.TienThanhToan;
                            break;
                        case EnumThang.Thang07:
                            totalView.DoanhThuThang7 += one.TienThanhToan;
                            break;
                        case EnumThang.Thang08:
                            totalView.DoanhThuThang8 += one.TienThanhToan;
                            break;
                        case EnumThang.Thang09:
                            totalView.DoanhThuThang9 += one.TienThanhToan;
                            break;
                        case EnumThang.Thang10:
                            totalView.DoanhThuThang10 += one.TienThanhToan;
                            break;
                        case EnumThang.Thang11:
                            totalView.DoanhThuThang11 += one.TienThanhToan;
                            break;
                        case EnumThang.Thang12:
                            totalView.DoanhThuThang12 += one.TienThanhToan;
                            break;
                        default:
                            break;
                    }
                }

                totals.Add(totalView);
                index++;
            }

            return totals;
        }
    }
}
