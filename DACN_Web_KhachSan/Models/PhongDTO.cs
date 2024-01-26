using Microsoft.Ajax.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DACN_Web_KhachSan.Models
{
    public class PhongDTO
    {
        public int Id { get; set; }
        public string TenPhong { get; set; }
        public int? MaLoai { get; set; }
        public int? NguoiLon { get; set; }
        public int? TreEm { get; set; }
        public int? DienTich { get; set; }
        public string LoaiGiuong { get; set; }
        public double? GiaTien { get; set; }
        public string TrangThai { get; set; }
        public string hinh {  get; set; }
    }
}