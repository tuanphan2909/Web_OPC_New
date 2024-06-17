using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using web4.Models;
using System.Web.Mvc;

namespace web4.Controllers
{
    public class DoanhThuChiNhanhController : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: DoanhThuChiNhanh
        public ActionResult Index()
        {
            return View();
        }

        public void connectSQL()  
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }

        public ActionResult DoanhThuChiNhanh(Account Acc)
        {
            DataSet ds = new DataSet();
            connectSQL();
            Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_BaoCaoDoanhThuCN_SAP]";
            ViewBag.ProcedureName = Pname;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_Tu_Ngay", Acc.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", Acc.To_date);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", Acc.Ma_DvCs_1);
                    sda.Fill(ds);

                }

            }
            return View(ds);
        }
        public ActionResult DoanhThuChiNhanh_Fill()
        {
            return View();
        }
        public ActionResult DoanhThuChiNhanh_Admin(Account Acc)
        {
            DataSet ds = new DataSet();
            connectSQL();
            Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_BaoCaoDoanhThuCN_SAP]";
            ViewBag.ProcedureName = Pname;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_Tu_Ngay", Acc.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", Acc.To_date);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", Acc.Ma_DvCs_1);
                    sda.Fill(ds);

                }

            }
            return View(ds);
        }
        public ActionResult DoanhThuChiNhanh_Admin_Fill()
        {
            return View();
        }
        public ActionResult DoanhThuChiNhanhTinhLuong(Account Acc)
        {

            DataSet ds = new DataSet();
            connectSQL();
            var MA_DVCS = "";
          
            if (Request.Cookies["UserName"].Value == "vpct-nhatanh")
            {
                 MA_DVCS = " ";
            }
            else
            {
                 MA_DVCS= Request.Cookies["MA_DVCS"].Value;
            }
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_BaoCaoDoanhThuCN_HCNS_SAP]";
            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_Tu_Ngay", Acc.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", Acc.To_date);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", MA_DVCS);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult DoanhThuChiNhanhTinhLuong_Fill()
        {
            return View();
        }
        public ActionResult DoanhThuChiNhanhKGam_Fill()
        {
            return View();
        }
        public ActionResult DoanhThuChiNhanhKGam(Account Acc)
        {
            DataSet ds = new DataSet();
            connectSQL();
            Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_KetQuaKinhDoanhCN_SAP]";
            ViewBag.ProcedureName = Pname;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_Tu_Ngay", Acc.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", Acc.To_date);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", Acc.Ma_DvCs_1);
                    sda.Fill(ds);

                }

            }
            return View(ds);
        }
        public ActionResult DoanhThuChiNhanhKGam_CN_Fill()
        {
            return View();
        }
        public ActionResult DoanhThuChiNhanhKGam_CN(Account Acc)
        {
            DataSet ds = new DataSet();
            connectSQL();
            Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_KetQuaKinhDoanhCN_SAP]";
            ViewBag.ProcedureName = Pname;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_Tu_Ngay", Acc.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", Acc.To_date);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", Acc.Ma_DvCs_1);
                    sda.Fill(ds);

                }

            }
            return View(ds);
        }
    }
}