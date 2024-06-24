using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace web4.Models
{
    public class TheoDoiGiaoHang
    {
        public string So_Ct { get; set; }
        public List<B30GDetail> Details { get; set; }
        public DateTime Ngay_Ct { get; set; }
        public string Ma_NVGH { get; set; }
        public string Ten_NVGH { get; set; }
        public string Ly_do { get; set; }
        public string Dvcs { get; set; }
        public string Ma_Dt { get; set; }
        public string Ten_Dt { get; set; }
        public string So_HD { get; set; }
        public string Ngay_HD { get; set; }
        public string NV_GiaoNhan { get; set; }
        public int Giao_HD { get; set; }
        public float Tien_HD { get; set; }
        public float Tien_Phai_Thu { get; set; }

        public string So_CT1 { get; set; }
        public string Noi_Dung { get; set; }
        public string Stt { get; set; }
        public int Da_Giao_hang { get; set; }
        public string Ten_NVPhuKho { get; set; }
        public string Ten_NVGH1 { get; set; }

        public string Han_TT { get; set; }
      
        public string Tien { get; set; }
        public string Tien1 { get; set; }
        public float Tien_CKTT { get; set; }
        public string Ngay_Den_Han { get; set; }
        public string Ngay_Nop_Tien { get; set; }

    }
}