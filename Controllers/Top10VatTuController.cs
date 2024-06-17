//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Web;
//using System.Web.Mvc;
//using StudentManagement.Models;
//using System.Data.SqlClient;
//using System.Data;
//namespace web4.Controllers
//{
//    public class Top10VatTuController : Controller
//    {
//        SqlConnection con = new SqlConnection();
//        SqlCommand sqlc = new SqlCommand();
//        SqlDataReader dt;
//        public void connectSQL()
//        {
//            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
//        }
//        // GET: Top10VatTu
//        public ActionResult MainBaoCao()
//        {
//            List<Top10DoanhThuItem> model = new List<Top10DoanhThuItem>();
//            connectSQL();
//            string Pname = "[usp_Top10DoanhThu_SAP]";
//            using (SqlCommand cmd = new SqlCommand(Pname, con))
//            {
//                cmd.CommandTimeout = 950;
//                cmd.Connection = con;
//                cmd.CommandType = CommandType.StoredProcedure;
//                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
//                {
//                    DataTable dt = new DataTable();
//                    sda.Fill(dt);

//                    foreach (DataRow row in dt.Rows)
//                    {
//                        Top10DoanhThuItem item = new Top10DoanhThuItem
//                        {
//                            MaVatTu = row.Field<string>("Ma_Vt"),
//                            TenVatTu = row.Field<string>("Ten_Vt"),
//                            DoanhThu = row.Field<string>("Doanh_Thu")
//                        };
//                        model.Add(item);
//                    }
//                }
//            }

//            return View("MainBaoCao", model);
//        }

//    }
//}