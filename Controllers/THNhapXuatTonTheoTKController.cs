using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using web4.Models;
using System.Web.Mvc;
using System.Net;
using System.Data.SqlClient;
using System.Data;
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
    public class THNhapXuatTonTheoTKController : Controller
    {
        // GET: THNhapXuatTonTheoTK
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }
        public ActionResult Index()
        {
            return View();
        }
        public List<BKHoaDonGiaoHang> LoadDmVt()
        {

            connectSQL();

            List<BKHoaDonGiaoHang> dataItems = new List<BKHoaDonGiaoHang>();

            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_PriceList_SAP]", connection))
                {
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;
                    using (SqlDataAdapter sda = new SqlDataAdapter(command))
                    {
                        DataSet ds = new DataSet();
                        sda.Fill(ds);

                        // Kiểm tra xem DataSet có bảng dữ liệu hay không
                        if (ds.Tables.Count > 0)
                        {
                            DataTable dt = ds.Tables[0];

                            foreach (DataRow row in dt.Rows)
                            {
                                BKHoaDonGiaoHang dataItem = new BKHoaDonGiaoHang
                                {
                                    Ma_Vt = row["Ma_Vt"].ToString(),
                                    Ten_Vt = row["Ten_Vt"].ToString(),
                                    Gia = row["Gia"].ToString()
                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }

            return dataItems;
        }

        public ActionResult TongNhapXuatTonTheoKy_Fill()
        {
            List<BKHoaDonGiaoHang> dmDlistVT = LoadDmVt();
            ViewBag.DataVT = dmDlistVT;
            return View();
        }
        public ActionResult TongNhapXuatTonTheoKy()
        {
            List<BKHoaDonGiaoHang> dmDlistVT = LoadDmVt();
            ViewBag.DataVT = dmDlistVT;
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            var Ma_Dvcs = Request.Cookies["MA_DVCS"].Value;
            DataSet ds = new DataSet();
            connectSQL();

            var Tk = Request.Cookies["Ma_Tk"].Value;


            var Ma_Dv = Request.Cookies["Ma_Dv"].Value;


            var dvcs = Request.Cookies["Ma_Dvcs"].Value;
            var Ma_Vt = Request.Cookies["Ma_Vt"].Value;
            //var Ma_Kho = Request.Cookies["Ma_Kho"].Value;
            //var Ma_DV = Request.Cookies["Ma_DV"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            if (dvcs == ""||Ma_Dvcs=="OPC_B1")
            {
                dvcs = "OPC";
                
            }
            string Pname = "[usp_TongHopNhapXuatTonTheoTK_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;
                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                con.Open();
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Ma_Dvcs", dvcs);
                    cmd.Parameters.AddWithValue("@_Tk", Tk);
                    cmd.Parameters.AddWithValue("@_Ma_Dv", Ma_Dv);
                    cmd.Parameters.AddWithValue("@_Ma_Vt", Ma_Vt);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult TongNhapXuatTonTheoKy_In()
        {
            List<BKHoaDonGiaoHang> dmDlistVT = LoadDmVt();
            ViewBag.DataVT = dmDlistVT;
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            var Ma_Dvcs = Request.Cookies["MA_DVCS"].Value;
            DataSet ds = new DataSet();
            connectSQL();

            var Tk = Request.Cookies["Ma_Tk"].Value;


            var Ma_Dv = Request.Cookies["Ma_Dv"].Value;


            var dvcs = Request.Cookies["Ma_Dvcs"].Value;
            var Ma_Vt = Request.Cookies["Ma_Vt"].Value;
            //var Ma_Kho = Request.Cookies["Ma_Kho"].Value;
            //var Ma_DV = Request.Cookies["Ma_DV"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            if (dvcs == "" || Ma_Dvcs == "OPC_B1")
            {
                dvcs = "OPC";

            }
            string Pname = "[usp_TongHopNhapXuatTonTheoTK_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;
                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                con.Open();
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Ma_Dvcs", dvcs);
                    cmd.Parameters.AddWithValue("@_Tk", Tk);
                    cmd.Parameters.AddWithValue("@_Ma_Dv", Ma_Dv);
                    cmd.Parameters.AddWithValue("@_Ma_Vt", Ma_Vt);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
    }
}