using System.Collections.Generic;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using SupportTuoiBeo.Data;
using SupportTuoiBeo.Data.Contexts;
using SupportTuoiBeo.Data.Entities;
using SupportTuoiBeo.Models;

namespace SupportTuoiBeo.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index(UserDetailsViewModel filterByMonth)
        {
            UserDetailsViewModel userDetailsViewModel = new UserDetailsViewModel();
            using (var db = new ApplicationDbContext())
            {
                if(filterByMonth.MonthSelected == EnumThang.All)
                {
                    IEnumerable<UserDetails> users = db.UserDetails;
                    userDetailsViewModel.UserDetails = users.Select(x => new UserDetailViewModel
                    {
                        Id = x.Id,
                        MaKH = x.MaKH,
                        Ngay = x.Ngay,
                        Thang = x.Thang,
                        TienThanhToan = x.TienThanhToan,
                        Tinh = x.Tinh
                    }).ToList();
                }
                else
                {
                    IEnumerable<UserDetails> users = db.UserDetails.Where(x => x.Thang == filterByMonth.MonthSelected);
                    userDetailsViewModel.UserDetails = users.Select(x => new UserDetailViewModel
                    {
                        Id = x.Id,
                        MaKH = x.MaKH,
                        Ngay = x.Ngay,
                        Thang = x.Thang,
                        TienThanhToan = x.TienThanhToan,
                        Tinh = x.Tinh
                    }).ToList();
                }
            }

            foreach(var user in userDetailsViewModel.UserDetails)
            {
                userDetailsViewModel.TotalMonthSelected += user.TienThanhToan;
            }

            return View(userDetailsViewModel);
        }

        public IActionResult Total()
        {
            var totals = new List<TotalViewModel>();
            var listUserDetail = new List<UserDetails>();

            using (var db = new ApplicationDbContext())
            {
                listUserDetail = db.UserDetails.ToList();
            }

            var listMaKH = listUserDetail.GroupBy(x => x.MaKH).Select(x => x.Key);
            var index = 1;

            foreach(var maKh in listMaKH)
            {
                var totalView = new TotalViewModel
                {
                    Id = index,
                    MaKH = maKh,
                    Tinh = listUserDetail.FirstOrDefault(x => x.MaKH == maKh).Tinh
                };
                var listOneKH = listUserDetail.Where(x => x.MaKH == maKh).ToList();
                foreach(var one in listOneKH)
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

            return View(totals);
        }
    }
}
