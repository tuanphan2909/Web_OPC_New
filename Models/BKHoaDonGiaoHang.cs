using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
namespace web4.Models
{
    public class BKHoaDonGiaoHang
    {
        public string So_Ct { get; set; }
        public string So_Ct_E { get; set; }
        public string Ma_Dt { get; set; }
        public string Ten_Dt { get; set; }
        public string Ma_Vt { get; set; }
        public string Ten_Vt { get; set; }
        public string Ma_TDV { get; set; }
        public string Ten_TDV { get; set; }
        public decimal Tong_Tien { get; set; }
        public string SelectedMaTDV { get; set; }
        public string From_date { get; set; }
        public string To_date { get; set; }
        public string Ma_DvCs { get; set; }
        public string UserName { get; set; }
        public string Ma_CbNv { get; set; }
        public string hoten { get; set; }
        public string Gia { get; set; }
        public string Ma_Dvcs { get; set; } // Định nghĩa thuộc tính Ma_Dvcs
    }

}
