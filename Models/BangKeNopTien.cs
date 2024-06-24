using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace web4.Models
{
    public class BangKeNopTien
    {
        public string So_BK { get; set; }
        public List<B30BKNT_Detail> Details { get; set; }
        public DateTime Ngay_Nop_Tien { get; set; }
        public string Dvcs { get; set; }
        public string Ma_TDV { get; set; }
        public string Ten_TDV { get; set; }
        public string So_HD { get; set; }
        public float Tong_Tien { get; set; }
        public string Ngay_HD { get; set; }
        public string Noi_Dung { get; set; }
        public string Stt { get; set; }
      

     
  
  

    }
}