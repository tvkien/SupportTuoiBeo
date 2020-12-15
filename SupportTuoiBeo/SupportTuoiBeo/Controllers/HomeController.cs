using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
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

            return View(userDetailsViewModel);
        }
    }
}
