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
    public class BaoCaoKetQuaKinhDoanhController : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: BaoCaoKetQuaKinhDoanh
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }
        public ActionResult BaoCaoKetQuaKinhDoanh_Fill()
        {
            //List<MauInChungTu> dmDlist = LoadDmDt("");
            //List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            //ViewBag.DataItems = dmDlist;
            //ViewBag.DataTDV = dmDlistTDV;
            return View();
        }
        public ActionResult BaoCaoKetQuaKinhDoanh()
        {
            //List<MauInChungTu> dmDlist = LoadDmDt("");
            //List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            string ma_dvcs = Request.Cookies["MA_DVCS"].Value;
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            var MaDt = Request.Cookies["Ma_Dt"] != null ? Request.Cookies["Ma_Dt"].Value : string.Empty;
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
            string Pname = "[usp_BaoCaoKetQuaKinhDoanh_SAP]";

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
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Ma_Dt", MaDt);
                    //cmd.Parameters.AddWithValue("@_Ma_CbNv", MaTDV);
                    cmd.Parameters.AddWithValue("@_Ma_DvCs", ma_dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public List<MauInChungTu> LoadDmDt(string Ma_dvcs)
        {
            connectSQL();

            Ma_dvcs = Request.Cookies["ma_dvcs"].Value;
            List<MauInChungTu> dataItems = new List<MauInChungTu>();
            string appendedString = Ma_dvcs == "OPC_B1" ? "_010203" : "_01"; // Dòng này cộng chuỗi dựa trên giá trị của Ma_dvcs
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DmDtTdv_SAP_MauIn]", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@_Ma_Dvcs", Ma_dvcs + appendedString);
                    command.CommandTimeout = 950;
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            MauInChungTu dataItem = new MauInChungTu
                            {
                                Ma_Dt = reader["ma_dt"].ToString(),
                                Ten_Dt = reader["ten_dt"].ToString(),
                                Dia_Chi = reader["Dia_Chi"].ToString(),
                                Dvcs = reader["Dvcs"].ToString(),
                                Dvcs1 = reader["Dvcs1"].ToString()
                            };
                            dataItems.Add(dataItem);
                        }
                    }
                }
            }

            return dataItems;
        }

        public List<BKHoaDonGiaoHang> LoadDmTDV()
        {
            string ma_dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : "";
            connectSQL();
            if (string.IsNullOrEmpty(ma_dvcs))
            {
                return null; // Trả về null nếu ma_dvcs rỗng
            }
            List<BKHoaDonGiaoHang> dataItems = new List<BKHoaDonGiaoHang>();

            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DanhSachTDV]", connection))
                {
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.AddWithValue("@_Ma_Dvcs", ma_dvcs);

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
                                    Ma_CbNv = row["Ma_CbNv"].ToString(),
                                    hoten = row["hoten"].ToString(),
                                    Ma_Dvcs = row["Ma_Dvcs"].ToString()
                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }

            return dataItems;
        }
    }
}