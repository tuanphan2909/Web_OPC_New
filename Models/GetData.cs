using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
namespace web4.Models
{
    public class GetData
    {
        //Mau So Ton No Chi Tiet Phai Thu
        public int Stt { get; set; }
        public string TenDt { get; set; }
        public string NgayCt { get; set; }
        public string SoCtEinv { get; set; }
        public string NgayDenHan { get; set; }
        public string GhiChu { get; set; }
        public decimal CongNoTT { get; set; }
        public decimal TienThue { get; set; }
        public decimal CongNo { get; set; }
        public decimal TotalCongNoTT { get; set; }
        public decimal TotalCongNoST { get; set; }
        public decimal TotalCongNo { get; set; }
        //Mau Thong Bao No QH
        public string  SoHD { get; set; }
        public string NgayXuat { get; set; }
        public decimal TienNo { get; set; }
        public string HanTT { get; set; }
        public int NgayQH { get; set; }
        //Mau Doi Chieu Cong No - Doi Chieu Doanh Thu Cong No
         public string So { get; set; }
        public string Ngay { get; set; }
        public string TienHD { get; set; }
        public string So1 { get; set; }
        public string Ngay1 { get; set; }
        public string SoTien { get; set; }
        public string So2 { get; set; }
        public string Ngay2 { get; set; }
        public decimal SoTien2 { get; set; }
        public string CKTT { get; set; }
        public string TongTien { get; set; }
       public string GhiChu1 { get; set; }


        public string So3 { get; set; }
        public string Ngay3 { get; set; }
        public decimal TienHD2 { get; set; }
        public string GhiChu2 { get; set; }
        //Mau Phieu Nhap XNTT
        public string NgayHD { get; set; }
        public string TienTT { get; set; }
        public string TienThanhToan { get; set; }
      
        

    }
}