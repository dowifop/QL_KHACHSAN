using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace QLKS.Models
{
    public class LoginViewModel
    {
        [Required(ErrorMessage = "Vui lòng nhập Tên Đăng Nhập.")]
        public string ma_kh { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập Mật Khẩu.")]
        [MinLength(6, ErrorMessage = "Mật khẩu phải chứa ít nhất 6 kí tự.")]
        public string mat_khau { get; set; }
    }
}