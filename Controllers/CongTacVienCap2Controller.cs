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

namespace web4.Controllers
{
    public class CongTacVienCap2Controller : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: BaoCaoTienVeCN


        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }

        // GET: CongTacVien
        public ActionResult index()
        {
            DataSet ds = new DataSet();
            connectSQL();

            string Ma_DvCs = Request.Cookies["MA_DVCS"].Value;
            //Acc.UserName = Request.Cookies["UserName"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "Danhsach_CTV_Cap2";


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
        public List<CTVCap2> LoadDmDt(string Ma_dvcs)
        {
            connectSQL();

            Ma_dvcs = Request.Cookies["ma_dvcs"].Value;
            List<CTVCap2> dataItems = new List<CTVCap2>();
            string appendedString = Ma_dvcs == "OPC_B1" ? "_010203" : "_01"; // Dòng này cộng chuỗi dựa trên giá trị của Ma_dvcs
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[DanhMucKhachHangCap2]", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@_Ma_Dvcs", Ma_dvcs);
                    command.CommandTimeout = 950;
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            CTVCap2 dataItem = new CTVCap2
                            {

                                Ma_Dt = reader["Ma_Dt"].ToString(),
                                Ten_Dt = reader["Ten_Dt"].ToString(),
                                Dvcs = reader["Dvcs"].ToString()

                            };
                            dataItems.Add(dataItem);
                        }
                    }
                }
            }

            return dataItems;
        }

        public List<CTVCap2> LoadDmVt()
        {
            connectSQL();

            //Ma_dvcs = Request.Cookies["ma_dvcs"].Value;
            List<CTVCap2> dataItems = new List<CTVCap2>();
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
                            CTVCap2 dataItem = new CTVCap2
                            {

                                Ma_vt = reader["Ma_Vt"].ToString(),
                                Ten_Vt = reader["Ten_Vt"].ToString()



                            };
                            dataItems.Add(dataItem);
                        }
                    }
                }
            }

            return dataItems;
        }
        public ActionResult InputCTVCap2()
        {
            List<CTVCap2> dmDlist = LoadDmDt("");
            List<CTVCap2> DmVt = LoadDmVt();

            ViewBag.DataItems = dmDlist;
            ViewBag.DataItems2 = DmVt;

            return View();
        }


        //public ActionResult SaveCtV(CTV CTV)
        //{

        //    connectSQL();
        //    using (SqlConnection connection = new SqlConnection(con.ConnectionString))
        //    {
        //        connection.Open();
        //        try
        //        {

        //            using (SqlCommand command = new SqlCommand("InsertB30CtvData", connection))
        //            {
        //                command.CommandType = CommandType.StoredProcedure;
        //                command.Parameters.AddWithValue("@_Ngay_Ct", CTV.Ngay_Ct);
        //                command.Parameters.AddWithValue("@_so_Ct", CTV.So_Ct);
        //                command.Parameters.AddWithValue("@_Dvcs", CTV.Dvcs);
        //                command.Parameters.AddWithValue("@_Ma_Dt", CTV.Ma_Dt);
        //                command.Parameters.AddWithValue("@_Ma_vt", CTV.Ma_vt);
        //                command.Parameters.AddWithValue("@_Ten_Vt", CTV.Ten_Vt);
        //                command.Parameters.AddWithValue("@_Han_Muc", CTV.Han_Muc);

        //                command.ExecuteNonQuery();

        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            if (con.State == ConnectionState.Open)
        //            {
        //                con.Close();
        //            }
        //            ex.Message.ToString();
        //        }
        //        return View();
        //    }

        //}
        public ActionResult SaveCtV(CTVCap2 CTV)
        {
            string result = "Error!";
            connectSQL();
            if (CTV != null && CTV.Details != null)
            {
                try
                {
                    var detailsTable = new DataTable();
                    detailsTable.Columns.Add("Ma_vt", typeof(string));
                    detailsTable.Columns.Add("Ten_Vt", typeof(string));
                    detailsTable.Columns.Add("So_Luong", typeof(int));
                    detailsTable.Columns.Add("Don_Gia", typeof(decimal));
                    detailsTable.Columns.Add("Thanh_Tien", typeof(decimal));


                    foreach (var detail in CTV.Details)
                    {
                        detailsTable.Rows.Add(detail.Ma_vt, detail.Ten_Vt, detail.So_Luong, detail.Don_Gia,detail.Thanh_Tien);
                    }

                    using (var connection = new SqlConnection(con.ConnectionString))
                    {
                        connection.Open();

                        using (var command = new SqlCommand("[InsertCongTacVienCap2_SAP]", connection))
                        {
                            command.CommandType = CommandType.StoredProcedure;

                            command.Parameters.AddWithValue("@_Ngay_Ct", CTV.Ngay_Ct);
                            command.Parameters.AddWithValue("@_so_Ct", CTV.So_Ct);
                            command.Parameters.AddWithValue("@_Ten_Dt", CTV.Ten_Dt);
                            command.Parameters.AddWithValue("@_Dvcs", CTV.Dvcs);
                            command.Parameters.AddWithValue("@_Ma_Dt", CTV.Ma_Dt);
                            command.Parameters.AddWithValue("@_Thue", CTV.Thue);
                            command.Parameters.AddWithValue("@_Tien_Thue", CTV.Tien_Thue);
                            command.Parameters.AddWithValue("@_Tien_Truoc_Thue", CTV.Tien_Truoc_Thue);
                            command.Parameters.AddWithValue("@_Tien_Sau_Thue", CTV.Tien_Sau_Thue);

                            // Pass details as a TVP parameter
                            var detailsParam = command.Parameters.AddWithValue("@_Details", detailsTable);
                            detailsParam.SqlDbType = SqlDbType.Structured;
                            detailsParam.TypeName = "B30CTVCap2DetailType"; // Replace with your actual type name

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
        public ActionResult UpdateCTV(CTVCap2 CTV)
        {
            string result = "Error!";
            connectSQL();
            if (CTV != null && CTV.Details != null)
            {
                try
                {
                    var detailsTable = new DataTable();
                    detailsTable.Columns.Add("Ma_vt", typeof(string));
                    detailsTable.Columns.Add("Ten_Vt", typeof(string));
                    detailsTable.Columns.Add("So_Luong", typeof(int));
                    detailsTable.Columns.Add("Don_Gia", typeof(decimal));
                    detailsTable.Columns.Add("Thanh_Tien", typeof(decimal));

                    foreach (var detail in CTV.Details)
                    {
                        detailsTable.Rows.Add(detail.Ma_vt, detail.Ten_Vt, detail.So_Luong,detail.Don_Gia,detail.Thanh_Tien);
                    }

                    using (var connection = new SqlConnection(con.ConnectionString))
                    {
                        connection.Open();

                        using (var command = new SqlCommand("UpdateCongTacVienCap2_SAP", connection))
                        {
                            command.CommandType = CommandType.StoredProcedure;

                            command.Parameters.AddWithValue("@_Ngay_Ct", CTV.Ngay_Ct);
                            command.Parameters.AddWithValue("@_so_Ct", CTV.So_Ct);
                            command.Parameters.AddWithValue("@_Ten_Dt", CTV.Ten_Dt);
                            command.Parameters.AddWithValue("@_Dvcs", CTV.Dvcs);
                            command.Parameters.AddWithValue("@_Ma_Dt", CTV.Ma_Dt);
                            command.Parameters.AddWithValue("@_Thue", CTV.Thue);
                            command.Parameters.AddWithValue("@_Tien_Thue", CTV.Tien_Thue);
                            command.Parameters.AddWithValue("@_Tien_Truoc_Thue", CTV.Tien_Truoc_Thue);
                            command.Parameters.AddWithValue("@_Tien_Sau_Thue", CTV.Tien_Sau_Thue);
                            command.Parameters.AddWithValue("@_CTVId", CTV.CTVId);

                            // Pass details as a TVP parameter
                            var detailsParam = command.Parameters.AddWithValue("@_Details", detailsTable);
                            detailsParam.SqlDbType = SqlDbType.Structured;
                            detailsParam.TypeName = "B30CTVCap2DetailType"; // Replace with your actual type name

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

            return Json(result, JsonRequestBehavior.AllowGet);
        }


        public ActionResult EditCTVCap2()
        {
            DataSet ds = new DataSet();
            connectSQL();
            List<CTVCap2> dmDlist = LoadDmDt("");
            List<CTVCap2> DmVt = LoadDmVt();

            ViewBag.DataItems = dmDlist;
            ViewBag.DataItems2 = DmVt;

            string Pname = "[EditHanMucCTVCap2]";
            string ctvId = Request.QueryString["CTVId"];
            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;


                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_CTVId", ctvId);
                    sda.Fill(ds);

                }
            }


            return View(ds);
        }

        public ActionResult CoppyCTVCap2()
        {
            DataSet ds = new DataSet();
            connectSQL();
            List<CTVCap2> dmDlist = LoadDmDt("");
            List<CTVCap2> DmVt = LoadDmVt();

            ViewBag.DataItems = dmDlist;
            ViewBag.DataItems2 = DmVt;

            string Pname = "[EditHanMucCTVCap2]";
            string ctvId = Request.QueryString["CTVId"];
            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;


                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_CTVId", ctvId);
                    sda.Fill(ds);

                }
            }


            return View(ds);
        }
    }


}