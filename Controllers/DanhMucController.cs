﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using web4.Models;
using System.Web.Mvc;


namespace web4.Controllers
{
    public class DanhMucController : Controller
    {

        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: BaoCaoTienVeCN
        public ActionResult Index()
        {
            return View();
        }
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }

        // GET: DanhMuc


        public ActionResult DanhMuc()
        {
            var username = Request.Cookies["UserName"].Value;
            ViewBag.Username = username;
            return View();
        }
        public ActionResult ViewDanhMucKH(Account Acc)
        {
            DataSet ds = new DataSet();
            connectSQL();
            Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_DmDt9CN_SAP]";
            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_ma_dvcs", Acc.Ma_DvCs);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }

        public ActionResult ViewDanhMucKHCN(Account Acc)
        {
            DataSet ds = new DataSet();
            connectSQL();
            Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_DmDt9CN_SAP]";
            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_ma_dvcs", Acc.Ma_DvCs_1);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult ViewDanhMucKH_Fill()
        {
            return View();
        }

        public ActionResult ViewDanhMucKH_FillCN()
        {
            return View();
        }

    }
}