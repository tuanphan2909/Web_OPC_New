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
    public class ChiTietCongNoPhaiTraController : Controller
    {
        // GET: ChiTietCongNoPhaiTra
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }
        public ActionResult ChiTietCongNoPhaiTra_Fill()
        {
            return View();
        }
        public ActionResult ChiTietCongNoPhaiTra()
        {
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_date"].Value;
            DataSet ds = new DataSet();

            string Pname = "[usp_ChiTietCongNoPhaiTra_SAP]";
            connectSQL();
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
                    sda.Fill(ds);
                }    
            }    
                return View(ds);
        }
    }
}