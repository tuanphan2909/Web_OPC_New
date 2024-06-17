using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace web4.Models
{
    public class SPTrungBay
    {
        public int Id { get; set; }
        public List<B20SPT_Detail> Details { get; set; }

        public DateTime Ngay_Ct { get; set; }
        public DateTime Ngay_Bat_Dau { get; set; }
        public DateTime Ngay_Ket_Thuc { get; set; }
        public string So_Ct { get; set; }
        public string Ma_Dt { get; set; }
        public string Ten_Dt { get; set; }
        public string Ma_SP { get; set; }
        public string Ten_SP { get; set; }
        public int Tien_TB { get; set; }
        public string Ma_vt { get; set; }
        public string Ten_Vt { get; set; }
        public string Dvt { get; set; }
        public int So_luong { get; set; }
        public string STT { get; set; }
        public string Dvcs { get; set; }
        public Boolean option_1 { get; set; }
        public Boolean option_2 { get; set; }
        public string Hinh_1 { get; set; }
        public string Hinh_2 { get; set; }

        public string Hinh_3 { get; set; }
        public string Ma_TDV { get; set; }



    }
}