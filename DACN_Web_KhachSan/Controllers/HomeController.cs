using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DACN_Web_KhachSan.Models;
using System.Configuration;
using System.IO;
using System.Web.UI.WebControls;
using Newtonsoft.Json.Linq;
using WebsiteVatlieuXayDung.Models;
using PagedList.Mvc;
using System.Net.NetworkInformation;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data.Linq;
using TimeZoneConverter;
using System.Data.Linq.SqlClient;

namespace DACN_Web_KhachSan.Controllers
{
    public class HomeController : Controller
    {
        DBKhachSanDataContext db = new DBKhachSanDataContext();

        public ActionResult Index()
        {
            var list = from p in db.PHONGs
                       join lp in db.LOAIPHONGs on p.MALOAI equals lp.MALOAI
                       group p by lp.MALOAI into grouped
                       select new PhongDTO
                       {
                           MaLoai = grouped.Key,
                           GiaTien = grouped.Min(p => p.GIA)

                       };
            ViewBag.Gia = list;
            return View(db.LOAIPHONGs);
        }

        #region xác thực người dùng
        public ActionResult DangNhap()
        {
            return View();
        }

        public ActionResult XuLyDangNhap(FormCollection f)
        {
            string dienThoai = f["txtDienThoai"];
            string matKhau = f["txtMatKhau"];
            bool isNhanVien = !string.IsNullOrEmpty(f["nhanvien"]);
            TAIKHOANKH kh = null;
            TAIKHOANNV nv = null;
            if (isNhanVien)
            {
                nv = db.TAIKHOANNVs.FirstOrDefault(t => t.DIENTHOAI == dienThoai && t.MATKHAU == matKhau);
                if (nv == null)
                {
                    ViewBag.kq = "Sai thông tin đăng nhập!";
                    return View("DangNhap");
                }
                Session["nv"] = nv;
                if (nv.LOAITAIKHOAN == "ql")
                {
                    return RedirectToAction("Index");
                }
                return RedirectToAction("Index");
            }
            else
            {
                kh = db.TAIKHOANKHs.FirstOrDefault(t => t.DIENTHOAI == dienThoai && t.MATKHAU == matKhau);
                if (kh == null)
                {
                    ViewBag.kq = "Sai thông tin đăng nhập!";
                    return View("DangNhap");
                }
                Session["kh"] = kh;
                return RedirectToAction("Index");
            }
        }

        public ActionResult DangXuat()
        {
            Session.Clear();
            return RedirectToAction("Index");
        }

        public ActionResult DangKy()
        {
            return View();
        }

        [HttpPost]
        public ActionResult DangKy(FormCollection f)
        {
            string hoTen = f["HoTen"];
            string gioiTinh = f["GioiTinh"];
            string email = f["Email"];
            string dienThoai = f["DienThoai"];
            string diaChi = f["DiaChi"];
            string matKhau = f["MatKhau"];
            string nlMatKhau = f["NlMatKhau"];

            if (hoTen == string.Empty || email == string.Empty || dienThoai == string.Empty || diaChi == string.Empty)
            {
                ViewBag.error = "Vui lòng nhập đầy đủ thông tin!";
                return View();
            }

            if (dienThoai.Length != 10 || dienThoai[0] != '0' || !dienThoai.All(char.IsDigit))
            {
                ViewBag.dt = "Số điện thoại không hợp lệ!";
                return View();
            }

            if (matKhau != nlMatKhau)
            {
                ViewBag.mk = "Mật khẩu không trùng khớp!";
                return View();
            }

            TAIKHOANKH tkx = db.TAIKHOANKHs.FirstOrDefault(t => t.DIENTHOAI == dienThoai);
            if (tkx != null)
            {
                ViewBag.dt = "Số điện thoại này đã được đăng ký!";
                return View();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    KHACHHANG kh = db.KHACHHANGs.FirstOrDefault(t => t.DIENTHOAI == dienThoai && t.EMAIL == email);
                    if (kh != null)
                    {
                        kh.HOTENKH = hoTen;
                        kh.GIOITINH = gioiTinh;
                        kh.DIENTHOAI = dienThoai;
                        kh.DIACHI = diaChi;
                        kh.EMAIL = email;
                    }
                    else
                    {
                        KHACHHANG khx = db.KHACHHANGs.FirstOrDefault(t => t.EMAIL == email);
                        if (khx != null)
                        {
                            ViewBag.email = "Email này đã được đăng ký!";
                            return View();
                        }

                        kh = new KHACHHANG();
                        kh.HOTENKH = hoTen;
                        kh.GIOITINH = gioiTinh;
                        kh.DIENTHOAI = dienThoai;
                        kh.DIACHI = diaChi;
                        kh.EMAIL = email;
                        kh.LOAIKH = "Du khách";
                        db.KHACHHANGs.InsertOnSubmit(kh);
                    }

                    TAIKHOANKH tk = new TAIKHOANKH();
                    tk.DIENTHOAI = dienThoai;

                    tk.MATKHAU = matKhau;
                    db.TAIKHOANKHs.InsertOnSubmit(tk);

                    db.SubmitChanges();
                }
                catch
                {
                    ViewBag.dt = "Đăng ký thất bại!";
                    return View();
                }

            }
            return RedirectToAction("DangNhap");
        }
        #endregion

        #region thông tin khách hàng
        public ActionResult TrangCaNhan()
        {
            TAIKHOANKH kh = (TAIKHOANKH)Session["kh"];
            if (kh == null)
            {
                return RedirectToAction("DangNhap");
            }
            return View();
        }

        public ActionResult DoiMatKhau()
        {
            return View();
        }

        [HttpPost]
        public ActionResult DoiMatKhau(FormCollection f)
        {
            string mkht = f["matkhauhientai"];
            string mkm = f["matkhaumoi"];
            string xnmk = f["xnmatkhaumoi"];

            TAIKHOANKH kh = (TAIKHOANKH)Session["kh"];
            if (kh == null)
            {
                return RedirectToAction("DangNhap");
            }
            TAIKHOANKH tk = db.TAIKHOANKHs.FirstOrDefault(t => t.DIENTHOAI == kh.DIENTHOAI);

            if (mkht != kh.MATKHAU)
            {
                ViewBag.er1 = "Mật khẩu không đúng!";
                return View();
            }
            if (mkm != xnmk)
            {
                ViewBag.er2 = "Mật khẩu không trùng khớp!";
                return View();
            }

            tk.MATKHAU = mkm;
            db.SubmitChanges();
            Session["kh"] = tk;
            return RedirectToAction("TrangCaNhan");
        }

        public ActionResult ChinhSuaThongTin()
        {
            if (Session["kh"] == null)
            {
                return RedirectToAction("DangNhap");
            }
            return View();
        }

        [HttpPost]
        public ActionResult ChinhSuaThongTin(FormCollection f)
        {
            string hoTen = f["hoten"];
            string gioiTinh = f["gioi"];
            string email = f["email"];
            string dienThoai = f["sdt"];
            string diaChi = f["diaChi"];

            TAIKHOANKH tk = (TAIKHOANKH)Session["kh"];
            if (tk == null)
            {
                return RedirectToAction("DangNhap");
            }

            if (hoTen == string.Empty || email == string.Empty || dienThoai == string.Empty || diaChi == string.Empty)
            {
                ViewBag.error = "Vui lòng nhập đầy đủ thông tin!";
                return View();
            }

            if (dienThoai.Length != 10 || dienThoai[0] != '0' || !dienThoai.All(char.IsDigit))
            {
                ViewBag.dt = "Số điện thoại không hợp lệ!";
                return View();
            }

            KHACHHANG khx = db.KHACHHANGs.FirstOrDefault(t => t.DIENTHOAI == dienThoai);
            if (khx != null)
            {
                if (khx.MAKH != tk.KHACHHANG.MAKH)
                {
                    ViewBag.dt = "Số điện thoại này đã tồn tại trong hệ thống!";
                    return View();
                }
            }

            khx = db.KHACHHANGs.FirstOrDefault(t => t.EMAIL == email);
            if (khx != null)
            {
                if (khx.MAKH != tk.KHACHHANG.MAKH)
                {
                    ViewBag.email = "Email này đã tồn tại trong hệ thống!";
                    return View();
                }
            }

            TAIKHOANKH tkkh = db.TAIKHOANKHs.FirstOrDefault(t => t.DIENTHOAI == tk.DIENTHOAI);
            db.TAIKHOANKHs.DeleteOnSubmit(tkkh);
            db.SubmitChanges();

            KHACHHANG kh = db.KHACHHANGs.FirstOrDefault(t => t.MAKH == tk.KHACHHANG.MAKH);
            kh.HOTENKH = hoTen;
            kh.GIOITINH = gioiTinh;
            kh.DIENTHOAI = dienThoai;
            kh.DIACHI = diaChi;
            kh.EMAIL = email;
            db.SubmitChanges();

            TAIKHOANKH tam = new TAIKHOANKH();
            tam.DIENTHOAI = dienThoai;
            tam.MATKHAU = tk.MATKHAU;
            db.TAIKHOANKHs.InsertOnSubmit(tam);
            db.SubmitChanges();

            Session["kh"] = tam;
            return RedirectToAction("TrangCaNhan");
        }
        #endregion

        #region QL Nhân Viên
        public ActionResult QLNhanVien()
        {
            if (!ktNhanVien(true))
            {
                return RedirectToAction("Index");
            }
            return View(db.NHANVIENs);
        }

        [HttpPost]
        public ActionResult QLNhanVien(FormCollection f)
        {
            string tim = f["tim"];
            ViewBag.tim = tim;
            var nhanviens = db.NHANVIENs.ToList();
            return View(nhanviens.Where(t => t.HOTENNV.Contains(tim) || t.GIOITINH.Contains(tim) || t.DIENTHOAI.Contains(tim) || t.DIACHI.Contains(tim) || t.EMAIL.Contains(tim) || t.CHUCVU.Contains(tim) || t.TRANGTHAI.Contains(tim) || t.LUONG.ToString().Contains(tim) || ((DateTime)t.NGAYVAOLAM).ToString("dd/MM/yyyy").Contains(tim)));
        }

        [HttpPost]
        public ActionResult XoaNV(int id)
        {
            try
            {
                HOADONDICHVU hd = db.HOADONDICHVUs.FirstOrDefault(t => t.MANV == id);
                if (hd != null)
                {
                    return Json(new { success = false, message = "Không thể xóa vì nhân viên này đã từng tạo hóa đơn dịch vụ!" });
                }

                NHANVIEN nv = db.NHANVIENs.FirstOrDefault(t => t.MANV == id);
                if (!string.IsNullOrEmpty(nv.HINH))
                {
                    var oldPath = Path.Combine(Server.MapPath("~/Image"), nv.HINH);
                    if (System.IO.File.Exists(oldPath))
                    {
                        System.IO.File.Delete(oldPath);
                    }
                }
                db.NHANVIENs.DeleteOnSubmit(nv);
                db.SubmitChanges();
                return Json(new { success = true, message = "Xóa nhân viên thành công!" });
            }
            catch
            {
                return Json(new { success = false, message = "Xóa nhân viên thất bại!" });
            }
        }

        public ActionResult GetNVById(int id)
        {
            NHANVIEN n = db.NHANVIENs.FirstOrDefault(t => t.MANV == id);
            NhanVienDTO nv = new NhanVienDTO();
            nv.MANV = n.MANV;
            nv.HOTENNV = n.HOTENNV;
            nv.GIOITINH = n.GIOITINH;
            nv.DIENTHOAI = n.DIENTHOAI;
            nv.DIACHI = n.DIACHI;
            nv.EMAIL = n.EMAIL;
            nv.NGAYVAOLAM = n.NGAYVAOLAM;
            nv.LUONG = n.LUONG;
            nv.TRANGTHAI = n.TRANGTHAI;
            nv.CHUCVU = n.CHUCVU;
            nv.HINH = n.HINH;
            return Json(new { data = nv }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult ThemNV(NHANVIEN nv, HttpPostedFileBase file)
        {
            try
            {
                NHANVIEN x = db.NHANVIENs.FirstOrDefault(t => t.DIENTHOAI == nv.DIENTHOAI);
                if (x != null)
                {
                    return Json(new { success = false, message = "Số điện thoại này đã có trên hệ thống!" });
                }

                x = db.NHANVIENs.FirstOrDefault(t => t.EMAIL == nv.EMAIL);
                if (x != null)
                {
                    return Json(new { success = false, message = "Email này đã có trên hệ thống!" });
                }
                var filename = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Image"), filename);
                if (System.IO.File.Exists(path))
                    return Json(new { success = false, message = "Hình ảnh đã tồn tại" });
                else
                    file.SaveAs(path);
                db.NHANVIENs.InsertOnSubmit(nv);
                db.SubmitChanges();
                return Json(new { success = true, message = "Thêm nhân viên thành công!" });
            }
            catch
            {
                return Json(new { success = false, message = "Thêm nhân viên thất bại!" });
            }
        }

        [HttpPost]
        public ActionResult SuaNV(NHANVIEN n, HttpPostedFileBase file)
        {
            try
            {
                NHANVIEN x = db.NHANVIENs.FirstOrDefault(t => t.DIENTHOAI == n.DIENTHOAI);
                if (x != null)
                {
                    if (x.MANV != n.MANV)
                    {
                        return Json(new { success = false, message = "Số điện thoại này đã có trên hệ thống!" });
                    }
                }

                x = db.NHANVIENs.FirstOrDefault(t => t.EMAIL == n.EMAIL);
                if (x != null)
                {
                    if (x.MANV != n.MANV)
                    {
                        return Json(new { success = false, message = "Email này đã có trên hệ thống!" });
                    }
                }

                NHANVIEN nv = db.NHANVIENs.FirstOrDefault(t => t.MANV == n.MANV);

                TAIKHOANNV tknv = db.TAIKHOANNVs.FirstOrDefault(t => t.DIENTHOAI == nv.DIENTHOAI);
                TAIKHOANKH tk = new TAIKHOANKH();
                if (tknv != null)
                {
                    tk.MATKHAU = tknv.MATKHAU;
                    db.TAIKHOANNVs.DeleteOnSubmit(tknv);
                    db.SubmitChanges();
                }

                if (file != null)
                {
                    var filename = Path.GetFileName(file.FileName);
                    var path = Path.Combine(Server.MapPath("~/Image"), filename);
                    if (System.IO.File.Exists(path))
                        return Json(new { success = false, message = "Hình ảnh đã tồn tại" });
                    else
                        file.SaveAs(path);

                    if (!string.IsNullOrEmpty(nv.HINH))
                    {
                        var oldPath = Path.Combine(Server.MapPath("~/Image"), nv.HINH);
                        if (System.IO.File.Exists(oldPath))
                        {
                            System.IO.File.Delete(oldPath);
                        }
                    }
                }

                nv.HOTENNV = n.HOTENNV;
                nv.GIOITINH = n.GIOITINH;
                nv.DIENTHOAI = n.DIENTHOAI;
                nv.DIACHI = n.DIACHI;
                nv.EMAIL = n.EMAIL;
                nv.NGAYVAOLAM = n.NGAYVAOLAM;
                nv.LUONG = n.LUONG;
                nv.TRANGTHAI = n.TRANGTHAI;
                nv.CHUCVU = n.CHUCVU;
                if (!string.IsNullOrEmpty(n.HINH))
                {
                    nv.HINH = n.HINH;
                }
                db.SubmitChanges();

                if (tknv != null)
                {
                    tk.DIENTHOAI = nv.DIENTHOAI;
                    db.TAIKHOANKHs.InsertOnSubmit(tk);
                    db.SubmitChanges();
                }

                return Json(new { success = true, message = "Sửa nhân viên thành công!" });
            }
            catch
            {
                return Json(new { success = false, message = "Sửa nhân viên thất bại!" });
            }
        }

        #endregion

        #region QL Khách Hàng
        public ActionResult QLKhachHang()
        {
            if (!ktNhanVien(true))
            {
                return RedirectToAction("Index");
            }
            return View(db.KHACHHANGs);
        }

        [HttpPost]
        public ActionResult QLKhachHang(FormCollection f)
        {
            string tim = f["tim"];
            ViewBag.tim = tim;
            return View(db.KHACHHANGs.Where(t => t.HOTENKH.Contains(tim) || t.GIOITINH.Contains(tim) || t.DIENTHOAI.Contains(tim) || t.DIACHI.Contains(tim) || t.EMAIL.Contains(tim) || t.LOAIKH.Contains(tim)));
        }

        public ActionResult GetKHById(int id)
        {
            KHACHHANG k = db.KHACHHANGs.FirstOrDefault(t => t.MAKH == id);
            KhachHangDTO kh = new KhachHangDTO();
            kh.MAKH = k.MAKH;
            kh.HOTENKH = k.HOTENKH;
            kh.DIACHI = k.DIACHI;
            kh.EMAIL = k.EMAIL;
            kh.DIENTHOAI = k.DIENTHOAI;
            kh.GIOITINH = k.GIOITINH;
            kh.LOAIKH = k.LOAIKH;
            return Json(new { data = kh }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult ThemKH(KHACHHANG kh)
        {
            try
            {
                KHACHHANG x = db.KHACHHANGs.FirstOrDefault(t => t.DIENTHOAI == kh.DIENTHOAI);
                if (x != null)
                {
                    return Json(new { success = false, message = "Số điện thoại này đã có trên hệ thống!" });
                }

                x = db.KHACHHANGs.FirstOrDefault(t => t.EMAIL == kh.EMAIL);
                if (x != null)
                {
                    return Json(new { success = false, message = "Email này đã có trên hệ thống!" });
                }

                db.KHACHHANGs.InsertOnSubmit(kh);
                db.SubmitChanges();
                return Json(new { success = true, message = "Thêm khách hàng thành công!" });
            }
            catch
            {
                return Json(new { success = false, message = "Thêm khách hàng thất bại!" });
            }
        }

        [HttpPost]
        public ActionResult SuaKH(KHACHHANG kh)
        {
            try
            {
                KHACHHANG x = db.KHACHHANGs.FirstOrDefault(t => t.DIENTHOAI == kh.DIENTHOAI);
                if (x != null)
                {
                    if (x.MAKH != kh.MAKH)
                    {
                        return Json(new { success = false, message = "Số điện thoại này đã có trên hệ thống!" });
                    }
                }

                x = db.KHACHHANGs.FirstOrDefault(t => t.EMAIL == kh.EMAIL);
                if (x != null)
                {
                    if (x.MAKH != kh.MAKH)
                    {
                        return Json(new { success = false, message = "Email này đã có trên hệ thống!" });
                    }
                }

                KHACHHANG k = db.KHACHHANGs.FirstOrDefault(t => t.MAKH == kh.MAKH);

                TAIKHOANKH tkkh = db.TAIKHOANKHs.FirstOrDefault(t => t.DIENTHOAI == k.DIENTHOAI);
                TAIKHOANKH tk = new TAIKHOANKH();
                if (tkkh != null)
                {
                    tk.MATKHAU = tkkh.MATKHAU;
                    db.TAIKHOANKHs.DeleteOnSubmit(tkkh);
                    db.SubmitChanges();
                }

                k.HOTENKH = kh.HOTENKH;
                k.DIACHI = kh.DIACHI;
                k.EMAIL = kh.EMAIL;
                k.GIOITINH = kh.GIOITINH;
                k.LOAIKH = kh.LOAIKH;
                k.DIENTHOAI = kh.DIENTHOAI;
                db.SubmitChanges();

                if (tkkh != null)
                {
                    tk.DIENTHOAI = kh.DIENTHOAI;
                    db.TAIKHOANKHs.InsertOnSubmit(tk);
                    db.SubmitChanges();
                }

                return Json(new { success = true, message = "Sửa khách hàng thành công!" });
            }
            catch
            {
                return Json(new { success = false, message = "Sửa khách hàng thất bại!" });
            }
        }

        [HttpPost]
        public ActionResult XoaKH(int id)
        {
            try
            {
                HOADON hd = db.HOADONs.FirstOrDefault(t => t.MAKH == id);
                if (hd != null)
                {
                    return Json(new { success = false, message = "Không thể xóa vì khách hàng này đã từng đặt phòng!" });
                }

                KHACHHANG kh = db.KHACHHANGs.FirstOrDefault(t => t.MAKH == id);
                db.KHACHHANGs.DeleteOnSubmit(kh);
                db.SubmitChanges();
                return Json(new { success = true, message = "Xóa khách hàng thành công!" });
            }
            catch
            {
                return Json(new { success = false, message = "Xóa khách hàng thất bại!" });
            }
        }
        #endregion

        #region QL Phòng
        public ActionResult QuanLyPhong()
        {
            if (!ktNhanVien(true))
            {
                return RedirectToAction("Index");
            }
            ViewBag.LoaiPhong = new SelectList(db.LOAIPHONGs.ToList().OrderBy(t => t.TENLOAI), "MALOAI", "TENLOAI");
            return View(db.PHONGs);
        }

        [HttpPost]
        public ActionResult QuanLyPhong(FormCollection f)
        {
            string tim = f["tim"];
            ViewBag.tim = tim;
            ViewBag.LoaiPhong = new SelectList(db.LOAIPHONGs.ToList().OrderBy(t => t.TENLOAI), "MALOAI", "TENLOAI");
            var phongs = db.PHONGs.ToList();
            return View(phongs.Where(t => t.TENPHONG.Contains(tim) || t.GIA.ToString().Contains(tim) || t.TRANGTHAI.Contains(tim) || t.LOAIPHONG.TENLOAI.Contains(tim) || t.DIENTICH.ToString().Contains(tim) || t.TREEM.ToString().Contains(tim) || t.NGUOILON.ToString().Contains(tim) || t.SOGIUONG.Contains(tim)));
        }

        [HttpPost]
        public ActionResult ThemPhong(PHONG p, HttpPostedFileBase file)
        {
            try
            {
                var filename = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Image"), filename);
                if (System.IO.File.Exists(path))
                    return Json(new { success = false, message = "Hình ảnh đã tồn tại" });
                else
                    file.SaveAs(path);
                db.PHONGs.InsertOnSubmit(p);
                db.SubmitChanges();
                return Json(new { success = true, message = "Thêm phòng mới thành công!" });
            }
            catch (Exception e)
            {
                return Json(new { success = false, message = e });
            }
        }

        public ActionResult GetPhongById(int id)
        {
            PHONG pHONG = db.PHONGs.FirstOrDefault(t => t.MAPHONG == id);
            PhongDTO p = new PhongDTO();
            p.Id = pHONG.MAPHONG;
            p.TenPhong = pHONG.TENPHONG;
            p.MaLoai = pHONG.MALOAI;
            p.NguoiLon = pHONG.NGUOILON;
            p.TreEm = pHONG.TREEM;
            p.DienTich = pHONG.DIENTICH;
            p.LoaiGiuong = pHONG.SOGIUONG;
            p.GiaTien = pHONG.GIA;
            p.TrangThai = pHONG.TRANGTHAI;
            return Json(new { data = p }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult SuaPhong(PHONG p, HttpPostedFileBase file)
        {
            PHONG pHONG = db.PHONGs.FirstOrDefault(t => t.MAPHONG == p.MAPHONG);

            if (file != null)
            {
                var filename = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Image"), filename);
                if (System.IO.File.Exists(path))
                    return Json(new { success = false, message = "Hình ảnh đã tồn tại" });
                else
                    file.SaveAs(path);

                if (!string.IsNullOrEmpty(pHONG.HINH))
                {
                    var oldPath = Path.Combine(Server.MapPath("~/Image"), pHONG.HINH);
                    if (System.IO.File.Exists(oldPath))
                    {
                        System.IO.File.Delete(oldPath);
                    }
                }
            }

            pHONG.MALOAI = p.MALOAI;
            pHONG.TENPHONG = p.TENPHONG;
            pHONG.NGUOILON = p.NGUOILON;
            pHONG.TREEM = p.TREEM;
            pHONG.DIENTICH = p.DIENTICH;
            pHONG.SOGIUONG = p.SOGIUONG;
            pHONG.GIA = p.GIA;
            pHONG.TRANGTHAI = p.TRANGTHAI;
            if (!string.IsNullOrEmpty(p.HINH))
            {
                pHONG.HINH = p.HINH;
            }

            db.SubmitChanges();
            return Json(new { success = true });
        }

        [HttpPost]
        public ActionResult XoaPhong(int id)
        {
            try
            {
                var l = db.DATPHONGs.Where(t => t.MAPHONG == id);
                if (l.Count() > 0)
                {
                    return Json(new { success = false, message = "Phòng này đang được đặt!" });
                }

                PHONG p = db.PHONGs.FirstOrDefault(t => t.MAPHONG == id);
                if (!string.IsNullOrEmpty(p.HINH))
                {
                    var oldPath = Path.Combine(Server.MapPath("~/Image"), p.HINH);
                    if (System.IO.File.Exists(oldPath))
                    {
                        System.IO.File.Delete(oldPath);
                    }
                }
                db.PHONGs.DeleteOnSubmit(p);
                db.SubmitChanges();
                return Json(new { success = true, message = "Xóa phòng thành công!" });
            }
            catch
            {
                return Json(new { success = false, message = "Lỗi xóa phòng!" });
            }
        }
        #endregion

        #region Sao Lưu & Phục Hồi
        public ActionResult SaoLuuPhucHoi()
        {
            return View();
        }
        [HttpPost]
        public ActionResult SaoLuu(string tentep)
        {
            try
            {
                string path = Server.MapPath("~/Restore/" + tentep);
                if (System.IO.File.Exists(path))
                {
                    System.IO.File.Delete(path);
                }
                SqlConnection con = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["masterConnectionString"].ConnectionString);
                if (con.State == ConnectionState.Closed)
                    con.Open();

                string query = "BACKUP DATABASE DBKHACHSAN TO DISK='" + path + "'";
                SqlCommand cmd = new SqlCommand(query, con);
                cmd.ExecuteNonQuery();

                if (con.State == ConnectionState.Open)
                    con.Close();
                return Json(new { success = true });
            }
            catch (Exception)
            {
                return Json(new { success = false });
            }
        }

        [HttpPost]
        public ActionResult PhucHoi(HttpPostedFileBase file)
        {
            try
            {
                var filename = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Restore"), filename);
                if (System.IO.File.Exists(path))
                {
                    System.IO.File.Delete(path);
                    file.SaveAs(path);
                }
                else
                    file.SaveAs(path);

                SqlConnection con = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["masterConnectionString"].ConnectionString);
                if (con.State == ConnectionState.Closed)
                    con.Open();

                string query = @"
                        use master
                        if DB_ID('DBKHACHSAN') is not null
                        BEGIN
	                        ALTER DATABASE DBKHACHSAN
	                        SET SINGLE_USER WITH ROLLBACK IMMEDIATE                        
	                        drop database DBKHACHSAN
                        END
                        RESTORE DATABASE DBKHACHSAN FROM DISK='" + path + @"' WITH RECOVERY
                        ALTER DATABASE DBKHACHSAN
                        SET MULTI_USER
                    ";
                SqlCommand cmd = new SqlCommand(query, con);
                cmd.ExecuteNonQuery();

                if (con.State == ConnectionState.Open)
                    con.Close();
                return Json(new { success = true });
            }
            catch (Exception)
            {
                return Json(new { success = false });
            }
        }
        #endregion

        public ActionResult LienHe()
        {
            return View();
        }

        public ActionResult DaGuiYeuCau()
        {
            return View();
        }

        public ActionResult LoaiPhong()
        {
            return View();
        }

        public ActionResult Booking()
        {
            var list = from p in db.PHONGs
                       join lp in db.LOAIPHONGs on p.MALOAI equals lp.MALOAI
                       group p by lp.MALOAI into grouped
                       select new PhongDTO
                       {
                           MaLoai = grouped.Key,
                           GiaTien = grouped.Min(p => p.GIA)

                       };
            ViewBag.Gia = list;
            return View(db.LOAIPHONGs);
        }

        [HttpPost]
        public ActionResult KetQuaTimPhong(FormCollection f, int? page)
        {
            if (page == null || page < 1)
                page = 1;
            int size = 5;

            string returnUrl = Request.UrlReferrer?.ToString();
            if (string.IsNullOrEmpty(f["checkin"]) || string.IsNullOrEmpty(f["checkout"]))
            {
                return Redirect(returnUrl);
            }

            DateTime checkin;
            DateTime.TryParseExact(f["checkin"], "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out checkin);
            DateTime checkout;
            DateTime.TryParseExact(f["checkout"], "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out checkout);
            int nguoilon = int.Parse(f["nguoilon"]);
            int treem = int.Parse(f["treem"]);

            ViewBag.checkin = f["checkin"];
            ViewBag.checkout = f["checkout"];
            ViewBag.nguoilon = nguoilon;
            ViewBag.treem = treem;

            var query = (from phong in db.PHONGs
                         join datPhong in db.DATPHONGs on phong.MAPHONG equals datPhong.MAPHONG into datPhongJoin
                         from datPhong in datPhongJoin.DefaultIfEmpty()
                         where phong.NGUOILON >= nguoilon
                               && phong.TREEM >= treem
                               && phong.TRANGTHAI == "Hoạt động"
                               && (
                                   (datPhong.NGAYNHANPHONG > checkin && datPhong.NGAYNHANPHONG > checkout)
                                   || (datPhong.NGAYTRAPHONG < checkin && datPhong.NGAYTRAPHONG < checkout)
                                   || datPhong.NGAYNHANPHONG == null
                                   || datPhong.NGAYTRAPHONG == null
                               )
                         select phong).Distinct();

            int tongpage = (int)Math.Ceiling((double)query.Count() / size);
            if (page > tongpage)
                page = tongpage;
            ViewBag.tongpage = tongpage;
            ViewBag.page = page;

            ViewBag.listphong = query.Skip(((int)page - 1) * size).Take(size).ToList();
            return View(db.LOAIPHONGs);
        }

        #region QL Dịch Vụ
        public ActionResult QLDichVu()
        {
            if (!ktNhanVien(true))
            {
                return RedirectToAction("Index");
            }
            return View(db.DICHVUs);
        }

        [HttpPost]
        public ActionResult QLDichVu(FormCollection f)
        {
            string tim = f["tim"];
            ViewBag.tim = tim;
            var dichvus = db.DICHVUs.ToList();
            return View(dichvus.Where(t => t.TENDV.Contains(tim) || t.DONGIA.ToString().Contains(tim) || t.MOTA.Contains(tim) || t.VITRI.ToString().Contains(tim)));
        }

        public ActionResult GetDVById(int id)
        {
            DICHVU d = db.DICHVUs.FirstOrDefault(t => t.MADV == id);
            DichVuDTO dv = new DichVuDTO();
            dv.MADV = d.MADV;
            dv.TENDV = d.TENDV;
            dv.DONGIA = d.DONGIA;
            dv.MOTA = d.MOTA;
            dv.VITRI = d.VITRI;
            return Json(new { data = dv }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult ThemDV(DICHVU dv, HttpPostedFileBase file)
        {
            try
            {
                var filename = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/Image"), filename);
                if (System.IO.File.Exists(path))
                    return Json(new { success = false, message = "Hình ảnh đã tồn tại" });
                else
                    file.SaveAs(path);

                db.DICHVUs.InsertOnSubmit(dv);
                db.SubmitChanges();
                return Json(new { success = true, message = "Thêm dịch vụ thành công!" });
            }
            catch
            {
                return Json(new { success = false, message = "Thêm dịch vụ thất bại!" });
            }
        }

        [HttpPost]
        public ActionResult SuaDV(DICHVU dv, HttpPostedFileBase file)
        {
            try
            {
                DICHVU d = db.DICHVUs.FirstOrDefault(t => t.MADV == dv.MADV);

                if (file != null)
                {
                    var filename = Path.GetFileName(file.FileName);
                    var path = Path.Combine(Server.MapPath("~/Image"), filename);
                    if (System.IO.File.Exists(path))
                        return Json(new { success = false, message = "Hình ảnh đã tồn tại" });
                    else
                        file.SaveAs(path);

                    if (!string.IsNullOrEmpty(d.HINH))
                    {
                        var oldPath = Path.Combine(Server.MapPath("~/Image"), d.HINH);
                        if (System.IO.File.Exists(oldPath))
                        {
                            System.IO.File.Delete(oldPath);
                        }
                    }
                }

                d.TENDV = dv.TENDV;
                d.DONGIA = dv.DONGIA;
                d.MOTA = dv.MOTA;
                d.VITRI = dv.VITRI;
                if (!string.IsNullOrEmpty(dv.HINH))
                {
                    d.HINH = dv.HINH;
                }

                db.SubmitChanges();

                return Json(new { success = true, message = "Sửa dịch vụ thành công!" });
            }
            catch
            {
                return Json(new { success = false, message = "Sửa dịch vụ thất bại!" });
            }
        }

        [HttpPost]
        public ActionResult XoaDV(int id)
        {
            try
            {
                var l = db.SUDUNGDICHVUs.Where(t => t.MADV == id);
                if (l.Count() > 0)
                {
                    return Json(new { success = false, message = "Không thể xóa dịch vụ vì dịch vụ này đã từng được đặt!" });
                }

                DICHVU dv = db.DICHVUs.FirstOrDefault(t => t.MADV == id);
                if (!string.IsNullOrEmpty(dv.HINH))
                {
                    var oldPath = Path.Combine(Server.MapPath("~/Image"), dv.HINH);
                    if (System.IO.File.Exists(oldPath))
                    {
                        System.IO.File.Delete(oldPath);
                    }
                }
                db.DICHVUs.DeleteOnSubmit(dv);
                db.SubmitChanges();
                return Json(new { success = true, message = "Xóa dịch vụ thành công!" });
            }
            catch
            {
                return Json(new { success = false, message = "Xóa dịch vụ thất bại!" });
            }
        }
        #endregion

        public ActionResult DichVuKhachSan()
        {
            return PartialView(db.DICHVUs);
        }

        [HttpPost]
        public ActionResult ChiTietPhong(FormCollection f, int id, int? page)
        {
            if (page == null || page < 1)
                page = 1;
            int size = 3;

            string returnUrl = Request.UrlReferrer?.ToString();
            if (string.IsNullOrEmpty(f["checkin"]) || string.IsNullOrEmpty(f["checkout"]))
            {
                return Redirect(returnUrl);
            }

            DateTime checkin;
            DateTime.TryParseExact(f["checkin"], "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out checkin);
            DateTime checkout;
            DateTime.TryParseExact(f["checkout"], "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out checkout);
            int nguoilon = int.Parse(f["nguoilon"]);
            int treem = int.Parse(f["treem"]);

            ViewBag.checkin = f["checkin"];
            ViewBag.checkout = f["checkout"];
            ViewBag.nguoilon = nguoilon;
            ViewBag.treem = treem;

            var query = (from phong in db.PHONGs
                         join loaiPhong in db.LOAIPHONGs on phong.MALOAI equals loaiPhong.MALOAI into loaiPhongGroup
                         from loaiPhong in loaiPhongGroup.DefaultIfEmpty()
                         join datPhong in db.DATPHONGs on phong.MAPHONG equals datPhong.MAPHONG into datPhongJoin
                         from datPhong in datPhongJoin.DefaultIfEmpty()
                         where phong.NGUOILON >= nguoilon
                               && phong.TREEM >= treem
                               && phong.MALOAI == id
                               && phong.TRANGTHAI == "Hoạt động"
                               && (
                                   (datPhong.NGAYNHANPHONG > checkin && datPhong.NGAYNHANPHONG > checkout)
                                   || (datPhong.NGAYTRAPHONG < checkin && datPhong.NGAYTRAPHONG < checkout)
                                   || datPhong.NGAYNHANPHONG == null
                                   || datPhong.NGAYTRAPHONG == null
                               )
                         select phong).Distinct();

            ViewBag.LoaiPhong = db.LOAIPHONGs.FirstOrDefault(t => t.MALOAI == id);
            var tiennghi = from tn in db.TIENNGHIs join cttn in db.CHITIETTIENNGHIs on tn.MATIENNGHI equals cttn.MATIENNGHI where cttn.MALOAI == id select tn;
            ViewBag.TienNghi = tiennghi;

            int tongpage = (int)Math.Ceiling((double)query.Count() / size);
            if (page > tongpage)
                page = tongpage;
            ViewBag.tongpage = tongpage;
            ViewBag.page = page;

            var listphong = query.Skip(((int)page - 1) * size).Take(size).ToList();
            return View(listphong);
        }

        public void LuuGioHang(List<DATPHONG> g)
        {
            Session["gh"] = g;
        }

        public List<DATPHONG> LayGioHang()
        {
            return (List<DATPHONG>)Session["gh"];
        }

        [HttpPost]
        public ActionResult chonPhong(DATPHONG dp)
        {
            int songay = ((DateTime)dp.NGAYTRAPHONG - (DateTime)dp.NGAYNHANPHONG).Days + 1;
            dp.TONGTIEN = songay * db.PHONGs.FirstOrDefault(t => t.MAPHONG == dp.MAPHONG).GIA;
            List<DATPHONG> g = LayGioHang();
            if (g == null)
            {
                g = new List<DATPHONG>();
                g.Add(dp);
            }
            else
            {
                DATPHONG dpc = g.FirstOrDefault(t => t.MAPHONG == dp.MAPHONG);
                if (dpc != null)
                {
                    return Json(new { success = false, message = "Bạn đã đặt phòng này!" });
                }
                g.Add(dp);
            }
            LuuGioHang(g);
            return Json(new { success = true, message = "Thêm phòng thành công!" });
        }

        public ActionResult GioHang()
        {
            ViewBag.dsp = db.PHONGs;
            return View();
        }

        public ActionResult xoaGH(int id)
        {
            List<DATPHONG> g = LayGioHang();
            DATPHONG xdp = g.FirstOrDefault(t => t.MAPHONG == id);
            g.Remove(xdp);
            LuuGioHang(g);
            return Json(new { success = true });
        }

        public ActionResult DienThongTinKH(int loaitt)
        {
            ViewBag.Key = loaitt;
            return View();
        }

        [HttpPost]
        public ActionResult DienThongTinKH(FormCollection f, int loaitt)
        {
            ViewBag.Key = loaitt;
            string dienThoai = f["sdt"];
            if (dienThoai.Length != 10 || dienThoai[0] != '0' || !dienThoai.All(char.IsDigit))
            {
                ViewBag.dt = "Số điện thoại không hợp lệ!";
                return View();
            }
            KHACHHANG kh = db.KHACHHANGs.FirstOrDefault(t => t.DIENTHOAI == dienThoai);
            if (kh == null)
            {
                kh = db.KHACHHANGs.FirstOrDefault(t => t.EMAIL == f["email"]);
                if (kh != null)
                {
                    ViewBag.email = "Email này đã được đăng ký!";
                    return View();
                }
                kh = new KHACHHANG();
                kh.HOTENKH = f["hoten"];
                kh.DIENTHOAI = dienThoai;
                kh.DIACHI = f["diaChi"];
                kh.EMAIL = f["email"];
                kh.GIOITINH = f["gioi"];
                kh.LOAIKH = "Du khách";

                db.KHACHHANGs.InsertOnSubmit(kh);
                db.SubmitChanges();
            }
            Session["kht"] = kh;
            if (int.Parse(f["loaitt"]) == 1)
            {
                return RedirectToAction("ThanhToanMoMo");
            }
            else
            {
                return RedirectToAction("ThanhToanVNPAY");
            }
        }


        public ActionResult ThongKe()
        {
            if (!ktNhanVien(false))
            {
                return RedirectToAction("Index");
            }
            return View();
        }

        [HttpPost]
        public ActionResult GetDoanhThuLoai(string bd, string kt)
        {
            List<DoanhThuLoaiDTO> result = layDoanhThu(bd, kt);

            return Json(new { data = result }, JsonRequestBehavior.AllowGet);
        }

        public List<DoanhThuLoaiDTO> layDoanhThu(string bd, string kt)
        {
            List<DoanhThuLoaiDTO> result = new List<DoanhThuLoaiDTO>();

            SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DBKHACHSANConnectionString"].ConnectionString);

            connection.Open();

            string sqlQuery = @"
                    SELECT LOAIPHONG.MALOAI, TENLOAI, SUM((DATEDIFF(day, NGAYNHANPHONG, NGAYTRAPHONG) + 1) * GIA) AS DOANHTHULOAI
                    FROM LOAIPHONG        
                    LEFT JOIN PHONG ON PHONG.MALOAI = LOAIPHONG.MALOAI
		            LEFT JOIN DATPHONG ON DATPHONG.MAPHONG = PHONG.MAPHONG
                    LEFT JOIN HOADON ON HOADON.MAHD = DATPHONG.MAHD
                    WHERE NGAYTHANHTOAN BETWEEN '" + bd + "' AND '" + kt + @"'
                    GROUP BY LOAIPHONG.MALOAI, TENLOAI
                ";

            SqlCommand command = new SqlCommand(sqlQuery, connection);

            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                DoanhThuLoaiDTO dtl = new DoanhThuLoaiDTO();
                dtl.TENLOAI = reader["TENLOAI"].ToString();
                dtl.DOANHTHULOAI = int.Parse(reader["DOANHTHULOAI"].ToString());
                result.Add(dtl);
            }
            return result;
        }

        public ActionResult ThanhToanMoMo()
        {
            int tongtien = 0;
            List<DATPHONG> g = LayGioHang();
            foreach (var item in g)
            {
                tongtien += (int)item.TONGTIEN;
            }

            string endpoint = "https://test-payment.momo.vn/gw_payment/transactionProcessor";
            string partnerCode = "MOMOOJOI20210710";
            string accessKey = "iPXneGmrJH0G8FOP";
            string serectkey = "sFcbSGRSJjwGxwhhcEktCHWYUuTuPNDB";
            string orderInfo = "Thanh toán hóa đơn đặt phòng";
            string returnUrl = "https://localhost:44317/Home/XacNhanThanhToan";
            string notifyurl = "https://bcf5-116-102-163-84.ngrok-free.app/Home/SavePayment/";

            string amount = tongtien + "";
            string orderid = DateTime.Now.Ticks.ToString();
            string requestId = DateTime.Now.Ticks.ToString();
            string extraData = "";

            string rawHash = "partnerCode=" +
                partnerCode + "&accessKey=" +
                accessKey + "&requestId=" +
                requestId + "&amount=" +
                amount + "&orderId=" +
                orderid + "&orderInfo=" +
                orderInfo + "&returnUrl=" +
                returnUrl + "&notifyUrl=" +
                notifyurl + "&extraData=" +
                extraData;

            MoMoSecurity crypto = new MoMoSecurity();
            //sign signature SHA256
            string signature = crypto.signSHA256(rawHash, serectkey);

            //build body json request
            JObject message = new JObject
            {
                { "partnerCode", partnerCode },
                { "accessKey", accessKey },
                { "requestId", requestId },
                { "amount", amount },
                { "orderId", orderid },
                { "orderInfo", orderInfo },
                { "returnUrl", returnUrl },
                { "notifyUrl", notifyurl },
                { "extraData", extraData },
                { "requestType", "captureMoMoWallet" },
                { "signature", signature }

            };

            string responseFromMomo = PaymentRequest.sendPaymentRequest(endpoint, message.ToString());

            JObject jmessage = JObject.Parse(responseFromMomo);

            return Redirect(jmessage.GetValue("payUrl").ToString());
        }

        public ActionResult XacNhanThanhToan(Result result)
        {
            if (result.errorCode == "0")
            {
                TAIKHOANKH tk = (TAIKHOANKH)Session["kh"];
                KHACHHANG kh;
                if (tk != null)
                {
                    kh = db.KHACHHANGs.FirstOrDefault(t => t.DIENTHOAI == tk.DIENTHOAI);
                }
                else
                {
                    kh = (KHACHHANG)Session["kht"];
                }

                HOADON hd = new HOADON();
                hd.MAHD = result.orderId;
                hd.MAKH = kh.MAKH;
                hd.NGAYTHANHTOAN = ngayHienTai();
                hd.TONGTIEN = double.Parse(result.amount);
                db.HOADONs.InsertOnSubmit(hd);

                List<DATPHONG> g = LayGioHang();
                foreach (var item in g)
                {
                    DATPHONG dp = new DATPHONG();
                    dp.NGAYDAT = item.NGAYDAT;
                    dp.NGAYNHANPHONG = item.NGAYNHANPHONG;
                    dp.NGAYTRAPHONG = item.NGAYTRAPHONG;
                    dp.TONGTIEN = item.TONGTIEN;
                    dp.MAPHONG = item.MAPHONG;
                    dp.MAHD = hd.MAHD;
                    db.DATPHONGs.InsertOnSubmit(dp);
                }
                db.SubmitChanges();

                Session.Remove("gh");
                db.SubmitChanges();
                ViewBag.KetQua = "Thanh toán thành công";
            }
            else
            {
                ViewBag.KetQua = "Thanh toán thất bại";
            }

            return View();
        }

        public string UrlPayment(int totalAmount, string orderId)
        {
            var urlPayment = "";
            TimeZoneInfo vietnamTimeZone = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time");
            DateTime vietnamDateTime = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, vietnamTimeZone);

            string vnp_Returnurl = ConfigurationManager.AppSettings["vnp_Returnurl"];
            string vnp_Url = ConfigurationManager.AppSettings["vnp_Url"];
            string vnp_TmnCode = ConfigurationManager.AppSettings["vnp_TmnCode"];
            string vnp_HashSecret = ConfigurationManager.AppSettings["vnp_HashSecret"];

            VnPayLibrary vnpay = new VnPayLibrary();

            vnpay.AddRequestData("vnp_Version", VnPayLibrary.VERSION);
            vnpay.AddRequestData("vnp_Command", "pay");
            vnpay.AddRequestData("vnp_TmnCode", vnp_TmnCode);
            vnpay.AddRequestData("vnp_Amount", (totalAmount * 100).ToString());

            vnpay.AddRequestData("vnp_BankCode", "VNBANK");

            vnpay.AddRequestData("vnp_CreateDate", vietnamDateTime.ToString("yyyyMMddHHmmss"));
            vnpay.AddRequestData("vnp_CurrCode", "VND");
            vnpay.AddRequestData("vnp_IpAddr", Utils.GetIpAddress());
            vnpay.AddRequestData("vnp_Locale", "vn");
            vnpay.AddRequestData("vnp_OrderInfo", "Thanh toan don hang:" + orderId);
            vnpay.AddRequestData("vnp_OrderType", "other");
            vnpay.AddRequestData("vnp_ReturnUrl", vnp_Returnurl);
            vnpay.AddRequestData("vnp_TxnRef", orderId);

            urlPayment = vnpay.CreateRequestUrl(vnp_Url, vnp_HashSecret);

            return urlPayment;
        }


        public ActionResult ThanhToanVNPAY()
        {
            int total = 0;
            List<DATPHONG> g = LayGioHang();
            foreach (var item in g)
            {
                total += (int)item.TONGTIEN;
            }
            string orderid = DateTime.Now.Ticks.ToString();
            return Redirect(UrlPayment(total, orderid).ToString());
        }

        public ActionResult XacNhanThanhToan2()
        {
            if (Request.QueryString.Count > 0)
            {
                string vnp_HashSecret = ConfigurationManager.AppSettings["vnp_HashSecret"];
                var vnpayData = Request.QueryString;
                VnPayLibrary vnpay = new VnPayLibrary();

                foreach (string s in vnpayData)
                {
                    if (!string.IsNullOrEmpty(s) && s.StartsWith("vnp_"))
                    {
                        vnpay.AddResponseData(s, vnpayData[s]);
                    }
                }
                string orderCode = Convert.ToString(vnpay.GetResponseData("vnp_TxnRef"));
                long vnpayTranId = Convert.ToInt64(vnpay.GetResponseData("vnp_TransactionNo"));
                string vnp_ResponseCode = vnpay.GetResponseData("vnp_ResponseCode");
                string vnp_TransactionStatus = vnpay.GetResponseData("vnp_TransactionStatus");
                String vnp_SecureHash = Request.QueryString["vnp_SecureHash"];
                String TerminalID = Request.QueryString["vnp_TmnCode"];
                long vnp_Amount = Convert.ToInt64(vnpay.GetResponseData("vnp_Amount")) / 100;
                String bankCode = Request.QueryString["vnp_BankCode"];

                bool checkSignature = vnpay.ValidateSignature(vnp_SecureHash, vnp_HashSecret);
                if (checkSignature)
                {
                    if (vnp_ResponseCode == "00" && vnp_TransactionStatus == "00")
                    {
                        TAIKHOANKH tk = (TAIKHOANKH)Session["kh"];
                        KHACHHANG kh;
                        if (tk != null)
                        {
                            kh = db.KHACHHANGs.FirstOrDefault(t => t.DIENTHOAI == tk.DIENTHOAI);
                        }
                        else
                        {
                            kh = (KHACHHANG)Session["kht"];
                        }

                        HOADON hd = new HOADON();
                        hd.MAHD = orderCode;
                        hd.MAKH = kh.MAKH;
                        hd.NGAYTHANHTOAN = ngayHienTai();
                        hd.TONGTIEN = vnp_Amount;
                        db.HOADONs.InsertOnSubmit(hd);

                        List<DATPHONG> g = LayGioHang();
                        foreach (var item in g)
                        {
                            DATPHONG dp = new DATPHONG();
                            dp.NGAYDAT = item.NGAYDAT;
                            dp.NGAYNHANPHONG = item.NGAYNHANPHONG;
                            dp.NGAYTRAPHONG = item.NGAYTRAPHONG;
                            dp.TONGTIEN = item.TONGTIEN;
                            dp.MAPHONG = item.MAPHONG;
                            dp.MAHD = hd.MAHD;
                            db.DATPHONGs.InsertOnSubmit(dp);
                        }
                        db.SubmitChanges();

                        Session.Remove("gh");
                        db.SubmitChanges();
                        ViewBag.KetQua = "Thanh toán thành công";

                    }
                    else
                    {
                        ViewBag.KetQua = "Thanh toán thất bại. Mã lỗi: " + vnp_ResponseCode;
                        //Thanh toan khong thanh cong.Ma loi: vnp_ResponseCode
                        //ViewBag.InnerText = "Có lỗi xảy ra trong quá trình xử lý.Mã lỗi: " + vnp_ResponseCode;
                        //log.InfoFormat("Thanh toan loi, OrderId={0}, VNPAY TranId={1},ResponseCode={2}", orderId, vnpayTranId, vnp_ResponseCode);
                    }
                    //displayTmnCode.InnerText = "Mã Website (Terminal ID):" + TerminalID;
                    //displayTxnRef.InnerText = "Mã giao dịch thanh toán:" + orderId.ToString();
                    //displayVnpayTranNo.InnerText = "Mã giao dịch tại VNPAY:" + vnpayTranId.ToString();
                    ViewBag.ThanhToanThanhCong = "Số tiền thanh toán (VND):" + vnp_Amount.ToString();
                    //displayBankCode.InnerText = "Ngân hàng thanh toán:" + bankCode;
                }
            }
            return View();
        }

        public bool ktNhanVien(bool ql)
        {
            TAIKHOANNV tk = (TAIKHOANNV)Session["nv"];
            if (tk == null)
                return false;
            if (ql)
            {
                if (tk.LOAITAIKHOAN != "ql")
                    return false;
            }
            return true;
        }

        public ActionResult XuatThongKe(string bd, string kt)
        {
            if (Session["nv"] != null)
            {
                TAIKHOANNV tk = (TAIKHOANNV)Session["nv"];
                string nguoilap = tk.NHANVIEN.HOTENNV;

                List<DoanhThuLoaiDTO> dt = layDoanhThu(bd, kt);

                DateTime tu = DateTime.Parse(bd);
                DateTime den = DateTime.Parse(kt);

                return ExportExcel(dt, nguoilap, tu.ToString("dd/MM/yyyy"), den.ToString("dd/MM/yyyy"));
            }
            return RedirectToAction("ThongKe");
        }

        public ActionResult ExportExcel(List<DoanhThuLoaiDTO> dt, string nguoilap, string tu, string den)
        {
            if (dt == null || (dt != null && dt.Count == 0))
            {
                return RedirectToAction("ThongKe");
            }

            for (int i = 1; i <= dt.Count; i++)
            {
                dt[i - 1].STT = i;
            }

            string templatePath = Server.MapPath("~/Models/ThongKe.xlsx");
            FileInfo templateFile = new FileInfo(templatePath);

            if (!templateFile.Exists)
            {
                return RedirectToAction("ThongKe");
            }

            var package = new ExcelPackage(templateFile);

            var worksheet = package.Workbook.Worksheets[0];

            worksheet.Cells["C12"].Value = ngayHienTai().ToString("dd/MM/yyyy");
            worksheet.Cells["E13"].Value = tu;
            worksheet.Cells["F13"].Value = den;

            int rowStart = 16;

            foreach (DoanhThuLoaiDTO row in dt)
            {
                worksheet.Cells[string.Format("C{0}", rowStart)].Value = row.STT;
                worksheet.Cells[string.Format("D{0}", rowStart)].Value = row.TENLOAI;
                worksheet.Cells[string.Format("E{0}", rowStart)].Value = row.DOANHTHULOAI;

                worksheet.Cells[string.Format("E{0}", rowStart)].Style.Numberformat.Format = "#,##0";

                ExcelRange cellC = worksheet.Cells[string.Format("C{0}", rowStart)];
                ExcelRange cellD = worksheet.Cells[string.Format("D{0}", rowStart)];
                ExcelRange cellE = worksheet.Cells[string.Format("E{0}", rowStart)];

                ApplyBorder(cellC);
                ApplyBorder(cellD);
                ApplyBorder(cellE);

                rowStart++;
            }
            worksheet.Cells[string.Format("E{0}", rowStart + 1)].Value = "Người lập";
            worksheet.Cells[string.Format("E{0}", rowStart + 2)].Value = nguoilap;

            ExcelRange lbNguoiLap = worksheet.Cells[string.Format("E{0}", rowStart + 1)];
            ExcelRange nguoiLap = worksheet.Cells[string.Format("E{0}", rowStart + 2)];

            CenterCell(lbNguoiLap);
            CenterCell(nguoiLap);

            var memoryStream = new MemoryStream();
            package.SaveAs(memoryStream);

            string fileName = "ThongKe.xlsx";
            string mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            memoryStream.Position = 0;

            return File(memoryStream, mimeType, fileName);

        }

        public void CenterCell(ExcelRange cell)
        {
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        public void ApplyBorder(ExcelRange cell)
        {
            cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }

        public ActionResult TaiFileBackup(string tentep)
        {
            string filePath = Server.MapPath("~/Restore/" + tentep);
            if (System.IO.File.Exists(filePath))
            {
                byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);

                return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, tentep);
            }
            else
            {
                return RedirectToAction("SaoLuuPhucHoi");
            }
        }

        public void layTimeZone(string timeZone)
        {
            Session["ClientTimeZone"] = TZConvert.IanaToWindows(timeZone);
        }

        public DateTime ngayHienTai()
        {
            TimeZoneInfo TimeZone = TimeZoneInfo.FindSystemTimeZoneById(Session["ClientTimeZone"].ToString());
            return TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZone);
        }
    }
}
