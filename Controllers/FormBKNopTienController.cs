using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using web4.Models;

namespace web4.Controllers
{
    public class FormBKNopTienController : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: BaoCaoTienVeCN


        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }

        // GET: TheoDoiGiaoHang

        public List<TheoDoiGiaoHang> LoadDmTDV()
        {
            string ma_dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : "";
            connectSQL();

            List<TheoDoiGiaoHang> dataItems = new List<TheoDoiGiaoHang>();

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
                                TheoDoiGiaoHang dataItem = new TheoDoiGiaoHang
                                {
                                    Ma_NVGH = row["Ma_CbNv"].ToString(),
                                    Ten_NVGH = row["hoten"].ToString(),
                                    Dvcs = row["Ma_Dvcs"].ToString()
                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }

            return dataItems;
        }


        public List<TheoDoiGiaoHang> LoadHD()
        {
            string ma_dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : "";
            string Ma_cbnv = Request.Cookies["Ma_NVGH"] != null ? Request.Cookies["Ma_NVGH"].Value : "";
            connectSQL();

            List<TheoDoiGiaoHang> dataItems = new List<TheoDoiGiaoHang>();

            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_RpCongNoTDV_SAP]", connection))
                {
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;

                   // command.Parameters.AddWithValue("@_Ma_Dvcs", ma_dvcs);
                    command.Parameters.AddWithValue("@_Ma_CbNv", Ma_cbnv);


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
                                TheoDoiGiaoHang dataItem = new TheoDoiGiaoHang
                                {
                                    So_CT1 = row["so_ct"].ToString(),
                                    Ngay_HD = row["Ngay_Ct"].ToString(),
                                    So_HD = row["So_HD"].ToString(),
                                    Ma_Dt = row["Ma_dt"].ToString(),
                                    Ten_Dt = row["Ten_Dt"].ToString(),                                  
                                    Tien_HD = float.Parse(row["Cong_No"].ToString()),
                                     Tien_CKTT = string.IsNullOrEmpty(row["Tien_CKTT"].ToString()) ? 0 : float.Parse(row["Tien_CKTT"].ToString()),
                                 Tien_Phai_Thu = string.IsNullOrEmpty(row["Tien_Phai_Thu"].ToString()) ? 0 : float.Parse(row["Tien_Phai_Thu"].ToString()),
                                 Ngay_Den_Han = row["Ngay_Den_Han"].ToString()

                            };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }

            return dataItems;
        }


        public List<TheoDoiGiaoHang> UpdateLoadHD()
        {
            string ma_dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : "";
            string Ma_cbnv = Request.Cookies["NV_GiaoHang"] != null ? Request.Cookies["NV_GiaoHang"].Value : "";
            connectSQL();

            List<TheoDoiGiaoHang> dataItems = new List<TheoDoiGiaoHang>();

            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_RpCongNoTDV_SAP]", connection))
                {
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.AddWithValue("@_Ma_Dvcs", ma_dvcs);
                    command.Parameters.AddWithValue("@_Ma_CbNv", Ma_cbnv);


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
                                TheoDoiGiaoHang dataItem = new TheoDoiGiaoHang
                                {

                                    So_Ct = row["so_ct"].ToString(),
                                    Ngay_HD = row["Ngay_Ct"].ToString(),
                                    So_HD = row["So_Hd"].ToString(),
                                    Ma_Dt = row["Ma_Dt"].ToString(),
                                    Ten_Dt = row["Ten_Dt"].ToString(),
                                    Tien_HD = float.Parse(row["Cong_No"].ToString())


                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }

            return dataItems;
        }

        public ActionResult InsertBangKe()
        {
            List<TheoDoiGiaoHang> dmDlistTDV = LoadDmTDV();

            ViewBag.DataTDV = dmDlistTDV;

            return View();
        }
        public ActionResult InsertBangKeLoadHD()
        {
            List<TheoDoiGiaoHang> dmDlistTDV = LoadDmTDV();
            List<TheoDoiGiaoHang> dmListHD = LoadHD();
            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataHD = dmListHD;
            return View();
        }
        public ActionResult SaveHD(BangKeNopTien BK)
        {
            BK.Dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : "";
            BK.Noi_Dung = Request.QueryString["Ly_Do"] != null ? Request.QueryString["Ly_Do"] : "";


            string result = "Error!";
            connectSQL();
            if (BK != null && BK.Details != null)
            {
                try
                {
                    var detailsTable = new DataTable();
                    detailsTable.Columns.Add("So_HD", typeof(int));
                    detailsTable.Columns.Add("So_CT", typeof(string));
                    detailsTable.Columns.Add("Ngay_HD", typeof(string));
                    detailsTable.Columns.Add("Ngay_Den_Han", typeof(string));
                    detailsTable.Columns.Add("Ma_Dt", typeof(int));
                    detailsTable.Columns.Add("Ten_Dt", typeof(string));
                    detailsTable.Columns.Add("Tien_HD", typeof(float));
                    detailsTable.Columns.Add("Tien_CKTT", typeof(float));
                    detailsTable.Columns.Add("Tien_CKTT", typeof(float));
                    detailsTable.Columns.Add("Tien_Phai_Thu", typeof(float));
                    detailsTable.Columns.Add("Noi_Dung", typeof(string));
                    
                    foreach (var detail in BK.Details)
                    {
                        detailsTable.Rows.Add(detail.So_HD,detail.So_CT, detail.Ngay_HD,detail.Ngay_Den_Han, detail.Ma_Dt, detail.Ten_Dt, detail.Tien_HD, detail.Tien_CKTT, detail.Tien_Phai_Thu,detail.Stt);
                    }

                    using (var connection = new SqlConnection(con.ConnectionString))
                    {
                        connection.Open();

                        using (var command = new SqlCommand("InsertBangKeNopTien", connection))
                        {
                            command.CommandType = CommandType.StoredProcedure;

                            command.Parameters.AddWithValue("@_Ngay_Nop_Tien", BK.Ngay_Nop_Tien);
                            command.Parameters.AddWithValue("@_So_BK", BK.So_BK);
                            command.Parameters.AddWithValue("@_Ma_TDV", BK.Ma_TDV);
                            command.Parameters.AddWithValue("@_Ten_TDV", BK.Ten_TDV);
                            command.Parameters.AddWithValue("@_Tong_Tien", BK.Tong_Tien);
                            command.Parameters.AddWithValue("@_Dvcs", BK.Dvcs);
                            command.Parameters.AddWithValue("@_Noi_Dung", BK.Noi_Dung);

                            // Pass details as a TVP parameter
                            var detailsParam = command.Parameters.AddWithValue("@_Details", detailsTable);
                            detailsParam.SqlDbType = SqlDbType.Structured;
                            detailsParam.TypeName = "B30BKNTdetailType"; // Replace with your actual type name

                            command.ExecuteNonQuery();

                            result = "Success! Đã Lưu";
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Handle exceptions appropriately
                    result = $"Error! {ex.Message}";
                }
            }

            return Json(result, JsonRequestBehavior.AllowGet);
        }
        public ActionResult Index()
        {
            DataSet ds = new DataSet();
            connectSQL();

            string Ma_DvCs = Request.Cookies["MA_DVCS"].Value;
            //Acc.UserName = Request.Cookies["UserName"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "DanhSachBangKeNopTien";


            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;


                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_Ma_Dvcs", Ma_DvCs);

                    sda.Fill(ds);

                }
            }


            return View(ds);
        }
        public ActionResult MauInBangKe()
        {
            DataSet ds = new DataSet();
            connectSQL();


            string Pname = "MauInGiaoHang";
            string Stt = Request.QueryString["STT"];
            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;


                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_Stt", Stt);
                    sda.Fill(ds);

                }
            }


            return View(ds);
        }
        public ActionResult UpdateBangKe()
        {
            List<TheoDoiGiaoHang> dmDlistTDV = LoadDmTDV();
            List<TheoDoiGiaoHang> dmListHD = LoadHD();
            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataHD = dmListHD;

            DataSet ds = new DataSet();
            connectSQL();

            string Pname = "[EditBangKeNopTien]";
            string Stt = Request.QueryString["STT"];

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;


                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_Stt", Stt);
                    sda.Fill(ds);

                }
            }


            return View(ds);
        }
        //public ActionResult SaveUpdate(BangKeNopTien BK)
        //{



        //    BK.Dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : "";
        //    BK.Ma_NVGH = Request.Cookies["Ma_NVGH"] != null ? Request.Cookies["Ma_NVGH"].Value : "";
        //    BK.Ten_NVGH = Request.Cookies["Ten_NVGH"] != null ? Request.Cookies["Ten_NVGH"].Value : "";
        //    BK.NV_GiaoNhan = Request.Cookies["NV_GiaoNhan"] != null ? Request.Cookies["NV_GiaoNhan"].Value : "";
        //    BK.Ly_do = Request.Cookies["Ly_Do"] != null ? Request.Cookies["Ly_Do"].Value : "";
        //    string result = "Error!";
        //    connectSQL();
        //    if (BK != null && BK.Details != null)
        //    {
        //        try
        //        {
        //            var detailsTable = new DataTable();
        //            detailsTable.Columns.Add("So_Hd", typeof(int));
        //            detailsTable.Columns.Add("Ngay_HD", typeof(string));
        //            detailsTable.Columns.Add("Ma_Dt", typeof(int));
        //            detailsTable.Columns.Add("Ten_Dt", typeof(string));
        //            detailsTable.Columns.Add("NV_GN", typeof(string));
        //            detailsTable.Columns.Add("Giao_HD", typeof(bool));
        //            detailsTable.Columns.Add("Tien", typeof(float));
        //            detailsTable.Columns.Add("Noi_Dung", typeof(string));
        //            detailsTable.Columns.Add("Chua_giao_hang", typeof(bool));
        //            foreach (var detail in BK.Details)
        //            {
        //                detailsTable.Rows.Add(detail.So_Hd, detail.Ngay_HD, detail.Ma_Dt, detail.Ten_Dt, detail.NV_GiaoNhan, detail.Giao_HD, detail.Tien_HD, detail.Noi_Dung, detail.Chua_giao_hang);
        //            }

        //            using (var connection = new SqlConnection(con.ConnectionString))
        //            {
        //                connection.Open();

        //                using (var command = new SqlCommand("UpdateBangKeNoptTien", connection))
        //                {
        //                    command.CommandType = CommandType.StoredProcedure;

        //                    command.Parameters.AddWithValue("@_Ngay_Ct", BK.Ngay_Ct);
        //                    command.Parameters.AddWithValue("@_NV_GiaoHang", BK.Ma_NVGH);
        //                    command.Parameters.AddWithValue("@_Ten_NVGiaoHang", BK.Ten_NVGH);
        //                    command.Parameters.AddWithValue("@_Ten_NVPhuKho", BK.NV_GiaoNhan);
        //                    command.Parameters.AddWithValue("@_Ly_Do", BK.Ly_do);
        //                    command.Parameters.AddWithValue("@_Stt", BK.Stt);


        //                    // Pass details as a TVP parameter
        //                    var detailsParam = command.Parameters.AddWithValue("@_Details", detailsTable);
        //                    detailsParam.SqlDbType = SqlDbType.Structured;
        //                    detailsParam.TypeName = "B30GdetailType"; // Replace with your actual type name

        //                    command.ExecuteNonQuery();

        //                    result = "Success! Đã Sửa";
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            // Handle exceptions appropriately
        //            result = $"Error! {ex.Message}";
        //        }

        //    }
        //    return Json(result, JsonRequestBehavior.AllowGet);
        //}




    }
}