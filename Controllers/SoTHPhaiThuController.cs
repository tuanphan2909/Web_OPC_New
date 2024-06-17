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
    public class SoTHPhaiThuController : Controller
    {
        // GET: SoTHPhaiThu
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
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
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }
        public ActionResult SoTongHopCongNoPhaiThu_Fill()
        {
            string ma_dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : string.Empty;
            List<MauInChungTu> dmDlist = LoadDmDt("");
            ViewBag.DataItems = dmDlist;
            return View();
        }
        public ActionResult SoTongHopCongNoPhaiThu()
        {
            List<MauInChungTu> dmDlist = LoadDmDt("");
          
            string ma_dvcs = Request.Cookies["MA_DVCS"].Value;
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            var MaDt = Request.Cookies["Ma_Dt"] != null ? Request.Cookies["Ma_Dt"].Value : string.Empty;
            var TK = Request.Cookies["Ma_TK"] != null ? Request.Cookies["Ma_TK"].Value : string.Empty;
            DataSet ds = new DataSet(); 
          
            ViewBag.DataItems = dmDlist;
            connectSQL();
           
            string Pname = "[usp_SoTongHopCongNoPhaiThu_SAP]";

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
                    cmd.Parameters.AddWithValue("@_Ma_Dt", MaDt);
                    cmd.Parameters.AddWithValue("@_Tk", TK);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult SoTongHopCongNoPhaiThu_In()
        {
            string ma_dvcs = Request.Cookies["MA_DVCS"].Value;
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            var MaDt = Request.Cookies["Ma_Dt"] != null ? Request.Cookies["Ma_Dt"].Value : string.Empty;
            var TK = Request.Cookies["Ma_TK"] != null ? Request.Cookies["Ma_TK"].Value : string.Empty;
            DataSet ds = new DataSet();

            connectSQL();

            string Pname = "[usp_SoTongHopCongNoPhaiThu_SAP]";

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
                    cmd.Parameters.AddWithValue("@_Ma_Dt", MaDt);
                    cmd.Parameters.AddWithValue("@_Tk", TK);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
    }
}