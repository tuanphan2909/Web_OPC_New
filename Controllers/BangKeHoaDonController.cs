using DevExpress.Office.Import.OpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using web4.Models;

namespace web4.Controllers
{
    public class BangKeHoaDonController : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: BangKeHoaDon
        public ActionResult Index()
        {
            return View();
        }
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }
        public ActionResult bangkehoadon(Account Acc)
        {
            DataSet ds = new DataSet();
            connectSQL();
            Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
            Acc.UserName = Request.Cookies["UserName"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_BKSaleOrder_SAP]";

            Acc.UserName = Request.Cookies["UserName"].Value;

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
                    cmd.Parameters.AddWithValue("@_username", Acc.UserName);
                    sda.Fill(ds);

                }
            }


            return View(ds);
           

        }
        public ActionResult bangkehoadon_Fill()
        {
            return View();
        }


        public ActionResult danhsachhoadon_SAP(Account Acc)
        {
            DataSet ds = new DataSet();
            connectSQL();
            List<MauInChungTu> dmDlist = LoadDmDt("");
            List<BKHoaDonGiaoHang> dmDlistVT = LoadDmVt();
            List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            ViewBag.DataItems = dmDlist;
            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataVT = dmDlistVT;
            var Ma_TDV = Request.Cookies["Ma_TDV"].Value;
            var Ma_Dt = Request.Cookies["Ma_Dt"].Value;
            var Ma_Vt = Request.Cookies["Ma_Vt"].Value;
            // Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
            //Acc.UserName = Request.Cookies["UserName"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_DanhSachHoaDon_SAP]";
            //var Ma_DvBh = Request.Cookies["Ma_DvBh"].Value;
            Acc.UserName = Response.Cookies["UserName"].Value;

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
                    cmd.Parameters.AddWithValue("@_Ma_Dt",Ma_Dt);
                    cmd.Parameters.AddWithValue("@_Ma_CbNv", Ma_TDV);
                    cmd.Parameters.AddWithValue("@_Tinh_Trang", Acc.Tinh_Trang);
                    cmd.Parameters.AddWithValue("@_username", Acc.UserName);
                    //cmd.Parameters.AddWithValue("@_Ma_DvBh", Ma_DvBh);

                    sda.Fill(ds);

                }
            }


            return View(ds);

        }
                           
        public ActionResult danhsachhoadon_SAP_CN(Account Acc)
        {
            DataSet ds = new DataSet();
            connectSQL();
            
            //Acc.UserName = Request.Cookies["UserName"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_DanhSachHoaDon_SAP]";
            //var Ma_DvBh = Request.Cookies["Ma_DvBh"].Value;
            Acc.UserName = Request.Cookies["UserName"].Value;

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
                    cmd.Parameters.AddWithValue("@_Ma_Dt", Acc.Ma_dt);
                    cmd.Parameters.AddWithValue("@_Tinh_Trang", Acc.Tinh_Trang);
                    cmd.Parameters.AddWithValue("@_username", Acc.UserName);
                    //cmd.Parameters.AddWithValue("@_Ma_DvBh", Ma_DvBh);
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
        public ActionResult DanhSachHoaDon_Fill()
        {
            List<MauInChungTu> dmDlist = LoadDmDt("");
            List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            List<BKHoaDonGiaoHang> dmDlistVT = LoadDmVt();
            ViewBag.DataItems = dmDlist;
            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataVT = dmDlistVT;
            return View();
        }
        public ActionResult DanhSachHoaDon_Fill_HCM()
        {
            return View();
        }
        public ActionResult bangkehoadon_Fill_HCM()
        {
            return View();
        }
    }
}