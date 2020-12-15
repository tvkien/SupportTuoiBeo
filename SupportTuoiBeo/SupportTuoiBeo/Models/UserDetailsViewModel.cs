using SupportTuoiBeo.Data;
using System;
using System.Collections.Generic;

namespace SupportTuoiBeo.Models
{
    public class UserDetailsViewModel
    {
        public List<UserDetailViewModel> UserDetails { get; set; }

        public EnumThang MonthSelected { get; set; }
    }
}