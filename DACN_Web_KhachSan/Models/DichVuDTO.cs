using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DACN_Web_KhachSan.Models
{
    public class DichVuDTO
    {
        public int MADV { get; set; }
        public string TENDV { get; set; }
        public double? DONGIA { get; set; }
        public string MOTA { get; set; }
        public int? VITRI { get; set; }
        public string HINH { get; set; }
    }
}