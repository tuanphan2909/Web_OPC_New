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
    public class BaoCaoDTGVLGController : Controller
    {
        // GET: BaoCaoDTGVLG
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }
        public ActionResult BaoCaoDTGVLG_Fill()
        {
            return View();
        }
        public ActionResult BaoCaoDTGVLG()
        {
          
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            var madvcs = Request.Cookies["MA_DVCS"].Value;
            DataSet ds = new DataSet();
            connectSQL();
            if (Request.Cookies["Ma_Dvcs_2"] != null)
            {
                madvcs = Request.Cookies["Ma_Dvcs_2"].Value;
            }
            else
            {
                madvcs = Request.Cookies["MA_DVCS"].Value;
            }
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_BaoCaoDTGVLG_SAP]";

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
                    cmd.Parameters.AddWithValue("@_ma_dvcs", madvcs);
                 
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
    }
}