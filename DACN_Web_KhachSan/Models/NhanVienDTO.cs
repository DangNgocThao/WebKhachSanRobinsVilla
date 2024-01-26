using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DACN_Web_KhachSan.Models
{
    public class NhanVienDTO
    {
        public int MANV { get; set; }
        public string HOTENNV { get; set; }
        public string GIOITINH { get; set; }
        public string DIENTHOAI { get; set; }
        public string  DIACHI { get; set; }
        public string EMAIL { get; set; }
        public DateTime? NGAYVAOLAM { get; set; }
        public double? LUONG { get; set; }
        public string TRANGTHAI { get; set; }
        public string CHUCVU { get; set; }
        public string  HINH { get; set; }
    }
}