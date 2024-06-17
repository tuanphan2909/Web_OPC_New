using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.SqlClient;
using System.Data;
using web4.Models;
using System.Drawing;
using OfficeOpenXml.Drawing;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using OfficeOpenXml.Table;
using Newtonsoft.Json;
using System.Globalization;
namespace web4.Controllers
{
    public class ChiTietSPTraThuongController : Controller
    {
        // GET: ChiTietSPTraThuong
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }
        public ActionResult ChiTietSPTraThuong_Fill()
        {
            return View();
        }
        public ActionResult ChiTietSPTraThuong()
        {
            //List<MauInChungTu> dmDlist = LoadDmDt("");
            //List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            string ma_dvcs;
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            if (Request.Cookies["Ma_Dvcs_2"] != null)
            {
                ma_dvcs = Request.Cookies["Ma_Dvcs_2"].Value;
            }
            else
            {
                ma_dvcs = Request.Cookies["MA_DVCS"].Value;
            }
            //var MaDt = Request.Cookies["Ma_Dt"] != null ? Request.Cookies["Ma_Dt"].Value : string.Empty;
            //var MaTDV = Request.Cookies["Ma_TDV"].Value;
            DataSet ds = new DataSet();
            //if (ma_dvcs == "OPC_B1")
            //{
            //    string ma_dvcsFirst3Chars = ma_dvcs == "OPC_B1" ? ma_dvcs.Substring(0, 3) : ma_dvcs;
            //    string first3Chars = ma_dvcsFirst3Chars.Substring(0, 3);
            //    ma_dvcs = first3Chars;
            //}
            //ViewBag.DataTDV = dmDlistTDV;
            //ViewBag.DataItems = dmDlist;
            connectSQL();
            //var SoCT = Request.Cookies["So_Ct"] != null ? Request.Cookies["So_Ct"].Value : "";
            //MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[Usp_DanhSachTrungBay_Detail]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                //MauIn.From_date = Request.Cookies["From_date"].Value;
                //MauIn.To_date = Request.Cookies["To_Date"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Ngay_Ct1", fromDate);
                    cmd.Parameters.AddWithValue("@_Ngay_Ct2", toDate);
                    //cmd.Parameters.AddWithValue("@_Ma_Dt", MaDt);
                    //cmd.Parameters.AddWithValue("@_Ma_CbNv", MaTDV);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
    }
}