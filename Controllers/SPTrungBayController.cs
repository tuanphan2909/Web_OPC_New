using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using web4.Models;
using System.Web.Mvc;
using System.Data;
using Newtonsoft.Json;
using System.Data.Entity.Validation;
using SixLabors.Fonts;
using SixLabors;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Drawing.Drawing2D;

namespace web4.Controllers
{
    public class SPTrungBayController : Controller
    {
        // GET: SPTrungBay
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: BaoCaoTienVeCN


        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }

        // GET: CongTacVien
        //public ActionResult index()
        //{
        //    DataSet ds = new DataSet();
        //    connectSQL();

        //    //string Ma_DvCs = Request.Cookies["MA_DVCS"].Value;
        //    //Acc.UserName = Request.Cookies["UserName"].Value;
        //    //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
        //    string Pname = "Danhsach_CTV";


        //    using (SqlCommand cmd = new SqlCommand(Pname, con))
        //    {
        //        cmd.CommandTimeout = 950;

        //        cmd.Connection = con;
        //        cmd.CommandType = CommandType.StoredProcedure;


        //        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
        //        {

        //            cmd.Parameters.AddWithValue("@_Ma_Dvcs", Ma_DvCs);

        //            sda.Fill(ds);

        //        }
        //    }


        //    return View(ds);
        //}
        public List<CTV> LoadDmDt(string Ma_dvcs)
        {
            connectSQL();

            Ma_dvcs = Request.Cookies["ma_dvcs"].Value;
            string Username = Request.Cookies["Username"].Value;
            List<CTV> dataItems = new List<CTV>();
            //string appendedString = Ma_dvcs == "OPC_B1" ? "_010203" : "_01"; // Dòng này cộng chuỗi dựa trên giá trị của Ma_dvcs
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DmDT_TDVSAP]", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@_Ma_Dvcs", Ma_dvcs + "_01");
                    command.Parameters.AddWithValue("@_Username", Username);
                    command.CommandTimeout = 950;
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            CTV dataItem = new CTV
                            {

                                Ma_Dt = reader["ma_dt"].ToString(),
                                Ten_Dt = reader["ten_dt"].ToString(),
                                Ma_TDV = reader["Ma_TDV"].ToString()
                            };
                            dataItems.Add(dataItem);
                        }
                    }
                }
            }

            return dataItems;
        }

        public List<CTV> LoadDmVt()
        {
            connectSQL();

            //Ma_dvcs = Request.Cookies["ma_dvcs"].Value;
            List<CTV> dataItems = new List<CTV>();
            //string appendedString = Ma_dvcs == "OPC_B1" ? "_010203" : "_01"; // Dòng này cộng chuỗi dựa trên giá trị của Ma_dvcs
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DanhMucVatTu]", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    // command.Parameters.AddWithValue("@_Ma_Dvcs", "OPC_HN_01");
                    command.CommandTimeout = 950;
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            CTV dataItem = new CTV
                            {

                                Ma_vt = reader["Ma_Vt"].ToString(),
                                Ten_Vt = reader["Ten_Vt"].ToString(),
                                Dvt = reader["Dvt"].ToString()

                            };
                            dataItems.Add(dataItem);
                        }
                    }
                }
            }

            return dataItems;
        }
        public ActionResult InsertSPTrungBay()
        {
            List<CTV> dmDlist = LoadDmDt("");
            List<CTV> DmVt = LoadDmVt();

            ViewBag.DataItems = dmDlist;
            ViewBag.DataItems2 = DmVt;

            return View();
        }
        public ActionResult SaveSPTB(SPTrungBay SPTB)
        {
            var Ma_dvcs = Request.Cookies["ma_dvcs"]?.Value;
            var Option_1 = Request.Cookies["Option_1"]?.Value;
            var Option_2 = Request.Cookies["Option_2"]?.Value;

            string result = "Error!";
            connectSQL();
            if (SPTB != null && SPTB.Details != null)
            {
                try
                {

                    byte[] img1 = !string.IsNullOrEmpty(SPTB.Hinh_1) ? Convert.FromBase64String(SPTB.Hinh_1.Split(',')[1]) : new byte[0];
                    byte[] img2 = !string.IsNullOrEmpty(SPTB.Hinh_2) ? Convert.FromBase64String(SPTB.Hinh_2.Split(',')[1]) : new byte[0];
                    byte[] img3 = !string.IsNullOrEmpty(SPTB.Hinh_3) ? Convert.FromBase64String(SPTB.Hinh_3.Split(',')[1]) : new byte[0];

                    var detailsTable = new DataTable();
                    detailsTable.Columns.Add("Ma_vt", typeof(int));
                    detailsTable.Columns.Add("Ten_Vt", typeof(string));
                    detailsTable.Columns.Add("Dvt", typeof(int));
                    detailsTable.Columns.Add("So_luong", typeof(int));

                    foreach (var detail in SPTB.Details)
                    {
                        detailsTable.Rows.Add(detail.Ma_Vt, detail.Ten_Vt, detail.Dvt, detail.So_luong);
                    }

                    using (var connection = new SqlConnection(con.ConnectionString))
                    {
                        connection.Open();

                        using (var command = new SqlCommand("InsertSPTB_SAP", connection))
                        {
                            command.CommandType = CommandType.StoredProcedure;

                            command.Parameters.AddWithValue("@_Ngay_Ct", SPTB.Ngay_Ct);
                            command.Parameters.AddWithValue("@_Ngay_Bat_Dau", SPTB.Ngay_Bat_Dau);
                            command.Parameters.AddWithValue("@_Ngay_Ket_Thuc", SPTB.Ngay_Ket_Thuc);
                            command.Parameters.AddWithValue("@_so_Ct", SPTB.So_Ct);
                            command.Parameters.AddWithValue("@_Ten_Dt", SPTB.Ten_Dt);
                            command.Parameters.AddWithValue("@_Tien_TB", SPTB.Tien_TB);
                            command.Parameters.AddWithValue("@_Ma_Dt", SPTB.Ma_Dt);
                            command.Parameters.AddWithValue("@_Ma_SP", SPTB.Ma_SP);
                            command.Parameters.AddWithValue("@_Ma_Dvcs", Ma_dvcs);
                            command.Parameters.AddWithValue("@_Ma_TDV", SPTB.Ma_TDV);
                            command.Parameters.AddWithValue("@_Ten_SP", SPTB.Ten_SP);
                            command.Parameters.AddWithValue("@_Option_1", Option_1);
                            command.Parameters.AddWithValue("@_Option_2", Option_2);
                            command.Parameters.AddWithValue("@_Hinh_1", (object)img1 ?? 0);
                            command.Parameters.AddWithValue("@_Hinh_2", (object)img2 ?? 0);
                            command.Parameters.AddWithValue("@_Hinh_3", (object)img3 ?? 0);

                            var detailsParam = command.Parameters.AddWithValue("@_Details", detailsTable);
                            detailsParam.SqlDbType = SqlDbType.Structured;
                            detailsParam.TypeName = "B30SPT_DetailType"; // Đổi tên thành tên loại thực tế của bạn

                            command.ExecuteNonQuery();

                            result = "Success! Đã Lưu";
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log lỗi và trả về chi tiết lỗi để gỡ lỗi
                    result = $"Error! {ex.Message}";
                }
            }

            if (result.StartsWith("Success"))
            {
                return Json(result, JsonRequestBehavior.AllowGet);
            }
            else
            {
                // Trả về lỗi với JSON để dễ dàng hiển thị trên client
                return new HttpStatusCodeResult(500, result);
            }
        }

       

        public ActionResult Index()
        {
            DataSet ds = new DataSet();
            connectSQL();

            string Ma_DvCs = Request.Cookies["MA_DVCS"].Value;
            string Username = Request.Cookies["UserName"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "DanhSachSPTrungBay";


            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;


                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_Ma_Dvcs", Ma_DvCs);
                    cmd.Parameters.AddWithValue("@_UserName", Username);

                    sda.Fill(ds);

                }
            }


            return View(ds);
        }
        public ActionResult UpdateSPTrungBay()

        {
            List<CTV> dmDlist = LoadDmDt("");
            List<CTV> DmVt = LoadDmVt();

            ViewBag.DataItems = dmDlist;
            ViewBag.DataItems2 = DmVt;
            DataSet ds = new DataSet();
            connectSQL();

            string Pname = "EditSPTrungBay";
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
        public ActionResult SaveUpdate(SPTrungBay TDGH)
        {



            TDGH.Dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : "";
            var Option_1 = Request.Cookies["Option_1"].Value;
            var Option_2 = Request.Cookies["Option_2"].Value;

            string result = "Error!";
            connectSQL();
            if (TDGH != null && TDGH.Details != null)
            {
                try
                {
                    byte[] img1 = !string.IsNullOrEmpty(TDGH.Hinh_1) ? Convert.FromBase64String(TDGH.Hinh_1.Split(',')[1]) : new byte[0];
                    byte[] img2 = !string.IsNullOrEmpty(TDGH.Hinh_2) ? Convert.FromBase64String(TDGH.Hinh_2.Split(',')[1]) : new byte[0];
                    byte[] img3 = !string.IsNullOrEmpty(TDGH.Hinh_3) ? Convert.FromBase64String(TDGH.Hinh_3.Split(',')[1]) : new byte[0];

                    //byte[] img1 = !string.IsNullOrEmpty(SPTB.Hinh_1) ? Convert.FromBase64String(SPTB.Hinh_1.Split(',')[1]) : new byte[0];
                    //byte[] img2 = !string.IsNullOrEmpty(SPTB.Hinh_2) ? Convert.FromBase64String(SPTB.Hinh_2.Split(',')[1]) : new byte[0];
                    //byte[] img3 = !string.IsNullOrEmpty(SPTB.Hinh_3) ? Convert.FromBase64String(SPTB.Hinh_3.Split(',')[1]) : new byte[0];

                    var detailsTable = new DataTable();
                    detailsTable.Columns.Add("Ma_Vt", typeof(int));
                    detailsTable.Columns.Add("Ten_Vt", typeof(string));
                    detailsTable.Columns.Add("Dvt", typeof(string));
                    detailsTable.Columns.Add("So_luong", typeof(int));

                    foreach (var detail in TDGH.Details)
                    {
                        detailsTable.Rows.Add(detail.Ma_Vt, detail.Ten_Vt, detail.Dvt, detail.So_luong);
                    }

                    using (var connection = new SqlConnection(con.ConnectionString))
                    {
                        connection.Open();

                        using (var command = new SqlCommand("UpdateSPTrungBay", connection))
                        {
                            command.CommandType = CommandType.StoredProcedure;

                            command.Parameters.AddWithValue("@_Ngay_Ct", TDGH.Ngay_Ct);
                            command.Parameters.AddWithValue("@_Ma_Dt", TDGH.Ma_Dt);
                            command.Parameters.AddWithValue("@_Ten_Dt", TDGH.Ten_Dt);
                            command.Parameters.AddWithValue("@_Ngay_Bat_Dau", TDGH.Ngay_Bat_Dau);
                            command.Parameters.AddWithValue("@_Ngay_Ket_Thuc", TDGH.Ngay_Ket_Thuc);
                            command.Parameters.AddWithValue("@_Ma_SP", TDGH.Ma_SP);
                            command.Parameters.AddWithValue("@_Ten_SP", TDGH.Ten_SP);
                            command.Parameters.AddWithValue("@_Tien_TB", TDGH.Tien_TB);
                            command.Parameters.AddWithValue("@_Stt", TDGH.STT);
                            command.Parameters.AddWithValue("@_Option_1", Option_1);
                            command.Parameters.AddWithValue("@_Option_2", Option_2);
                            command.Parameters.AddWithValue("@_Hinh_1", (object)img1 ?? DBNull.Value);
                            command.Parameters.AddWithValue("@_Hinh_2", (object)img2 ?? DBNull.Value);
                            command.Parameters.AddWithValue("@_Hinh_3", (object)img3 ?? DBNull.Value);

                            // Pass details as a TVP parameter
                            var detailsParam = command.Parameters.AddWithValue("@_Details", detailsTable);
                            detailsParam.SqlDbType = SqlDbType.Structured;
                            detailsParam.TypeName = "B30SPT_DetailType"; // Replace with your actual type name

                            command.ExecuteNonQuery();

                            result = "Success! Đã Sửa";
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Handle exceptions appropriately
                    result = $"Error! {ex.Message}";
                }

            }
            if (result.StartsWith("Success"))
            {
                return Json(result, JsonRequestBehavior.AllowGet);
            }
            else
            {
                // Trả về lỗi với JSON để dễ dàng hiển thị trên client
                return new HttpStatusCodeResult(500, result);
            }
        }

        public ActionResult Delete(SPTrungBay SP)
        {

           
            string result = "Error!";
            connectSQL();
            try
            {
                using (var connection = new SqlConnection(con.ConnectionString))
                {
                    connection.Open();

                    using (var command = new SqlCommand("DELETE_SPTrungBay", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@_STT", SP.STT);



                        command.ExecuteNonQuery();

                        result = "Success! Đã Xoá";
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions appropriately
                result = $"Error! {ex.Message}";
            }

            
            return Json(result, JsonRequestBehavior.AllowGet);
        }

    }

}
