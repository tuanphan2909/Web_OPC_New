using System;
using System.Web.Mvc;
using System.Data;
using System.Data.SqlClient;
using System.Web.Caching;
using web4.Models;

namespace web4.Controllers
{
    public class BaoCaoHopDongController : Controller
    {
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

        public ActionResult BaoCaoHopDongMar(Account Acc)
        {
            // Trước tiên, kiểm tra xem dữ liệu đã tồn tại trong Cache chưa
            DataSet ds = new DataSet();

            ds = new DataSet();
                connectSQL();
                string Pname = "[usp_BaoCaoHopDongMAR_SAP]";
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
                        sda.Fill(ds);
                    }
                }
            return View(ds);
        }

        public ActionResult BaoCaoHopDongMar_Fill()
        {
            return View();
        }

        public ActionResult BaoCaoChiTietHD_Fill()
        {
            return View();
        }

        public ActionResult BaoCaoChiTietHD(Account Acc)
        {
            // Tương tự, áp dụng caching cho phương thức này nếu cần                                                                       
            //string cacheKey = $"BaoCaoChiTietHD_{Acc.From_date}_{Acc.To_date}_{Acc.Ma_DvCs_1}";
            DataSet ds = new DataSet();

            connectSQL();
            // Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
            //Acc.UserName = Request.Cookies["UserName"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_BaoCaoDTGamMAR_SAP]";
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
                    sda.Fill(ds);

                }
            }

            return View(ds);
        }
    }
}
