using System;
using System.Web.Mvc;
using System.Data;
using System.Data.SqlClient;
using System.Web.Caching;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using web4.Models;
namespace web4.Controllers
{
    public class TheoDoiHDThauController : Controller
    {
        // GET: BaoCao
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        public ActionResult Index()
        {
            return View();
        }
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
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
        public ActionResult TheoDoiHopDongThau_Fill()
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"] != null ? Request.Cookies["Ma_dvcs"].Value : string.Empty;
            if (string.IsNullOrEmpty(ma_dvcs))
            {
                return View(); // Trả về null nếu ma_dvcs rỗng
            }

            // Gọi LoadDmHD với Ma_TDV để lấy dữ liệu đã lọc theo Ma_TDV
            List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            List<MauInChungTu> dmDlist = LoadDmDt("");
            //var distinctDataTDV = dmDList
            //    .GroupBy(x => x.Ten_TDV)
            //    .Select(x => x.First())
            //    .ToList();

            //// var distinctDataItems = dmDList
            ////.GroupBy(x => x.So_Ct_E)
            ////.Select(x => x.First())
            //.ToList();


            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataItems = dmDlist;


            DataSet ds = new DataSet();
            connectSQL();
            string Pname = "[usp_DanhSachTDV]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;
                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                    sda.Fill(ds);
                }
            }
            return View();
        }

        public ActionResult TheoDoiHopDongThau()
        {
            DataSet ds = new DataSet();
            connectSQL();
            List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            List<MauInChungTu> dmDlist = LoadDmDt("");
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_TheoDoiHopDongThau_SAP]";
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;

            var Dvcs = Request.Cookies["MA_DVCS"].Value==""? Request.Cookies["Dvcs3"].Value : Request.Cookies["MA_DVCS"].Value;
            var MaTDV = Request.Cookies["Ma_CbNv"] != null ? Request.Cookies["Ma_CbNv"].Value : string.Empty;
            var MaDt = Request.Cookies["Ma_DT"] != null ? Request.Cookies["Ma_DT"].Value : string.Empty;
            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataItems = dmDlist;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Ma_Dt", MaDt);
                    cmd.Parameters.AddWithValue("@_Ma_CbNv", MaTDV);
                    cmd.Parameters.AddWithValue("@_Ma_DvCs", Dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);

        }
        public ActionResult TheoDoiHopDongThau_All()
        {
            DataSet ds = new DataSet();
            connectSQL();
            List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            List<MauInChungTu> dmDlist = LoadDmDt("");
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_TheoDoiHopDongThau_SAP]";
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;

            var Dvcs = Request.Cookies["MA_DVCS"].Value == "" ? Request.Cookies["Dvcs3"].Value : Request.Cookies["MA_DVCS"].Value;
            var MaTDV = Request.Cookies["Ma_TDV"] != null ? Request.Cookies["Ma_TDV"].Value : string.Empty;
            var MaDt = Request.Cookies["Ma_DT"] != null ? Request.Cookies["Ma_DT"].Value : string.Empty;
            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataItems = dmDlist;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Ma_Dt", MaDt);
                    cmd.Parameters.AddWithValue("@_Ma_CbNv", MaTDV);
                    cmd.Parameters.AddWithValue("@_Ma_DvCs", Dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);

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