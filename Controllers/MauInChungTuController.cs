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
using System.Web.Http;
using System.Runtime.Caching;

namespace web4.Controllers
{
    public class MauInChungTuController : Controller
    {

        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: BaoCaoTienVeCN

        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }

        // GET: DanhMuc



        public ActionResult Index(MauInChungTu MauIn)
        {
            DataSet ds = new DataSet();
            connectSQL();

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_MauInChungTu_SAP]";
            Response.Cookies["From_date"].Value = MauIn.From_date.ToString();
            Response.Cookies["To_Date"].Value = MauIn.To_date.ToString();
            MauIn.UserName = Request.Cookies["UserName"].Value;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_username", MauIn.UserName);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }

        public ActionResult MauInNLCB(MauInChungTu MauIn)
        {
            DataSet ds = new DataSet();
            connectSQL();

            MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_MauInChungTuDetail_SAP]";



            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                MauIn.From_date = Request.Cookies["From_date"].Value;
                MauIn.To_date = Request.Cookies["To_Date"].Value;
                MauIn.UserName = Request.Cookies["UserName"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_So_Ct", MauIn.So_Ct);
                    cmd.Parameters.AddWithValue("@_username", MauIn.UserName);


                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult MauInNLCB_Fill()
        {
            return View();
        }

        public List<MauInChungTu> LoadDmDt(string Ma_dvcs)
        {
            connectSQL();

            Ma_dvcs = Request.Cookies["ma_dvcs"].Value;
            List<MauInChungTu> dataItems = new List<MauInChungTu>();
            string appendedString = Ma_dvcs == "OPC_B1" ? "_0120" : "_01"; // Dòng này cộng chuỗi dựa trên giá trị của Ma_dvcs
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

        public ActionResult PhieuNhapXNTT_Fill()
        {
            List<MauInChungTu> dmDlist = LoadDmDt("");

            ViewBag.DataItems = dmDlist;
            return View();
        }
        public ActionResult PhieuNhapXNTT_All()
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            var MaDt = Request.Cookies["Ma_DT"] != null ? Request.Cookies["Ma_DT"].Value : string.Empty;
            DataSet ds = new DataSet();
            connectSQL();
            var SoCT = Request.Cookies["So_Ct"] != null ? Request.Cookies["So_Ct"].Value : "";
            //MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_XacNhanCKTT_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                //MauIn.From_date = Request.Cookies["From_date"].Value;
                //MauIn.To_date = Request.Cookies["To_Date"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {


                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Ma_dt", MaDt);
                    cmd.Parameters.AddWithValue("@_so_Ct", SoCT);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult PhieuNhapXNTT_Index(MauInChungTu MauIn)
        {
            DataSet ds = new DataSet();
            connectSQL();
            List<MauInChungTu> dmDlist = LoadDmDt("");
            var dvcs = Request.Cookies["MA_DVCS"].Value;
            ViewBag.DataItems = dmDlist;
            string Pname = "[usp_XacNhanCKTT_SAP]";
            //var fromDate = Request.Cookies["From_date"].Value;
            //var toDate = Request.Cookies["To_Date"].Value;

            var MaDt = Request.Cookies["Ma_DT"].Value;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;
                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;


                cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                cmd.Parameters.AddWithValue("@_Ma_Dt", MaDt);
                cmd.Parameters.AddWithValue("@_ma_dvcs", dvcs);
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    sda.Fill(ds);
                }
            }

            return View(ds);
        }

        public ActionResult PhieuNhapXNTT(MauInChungTu MauIn)
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            //var fromDate = MauIn.From_date =="" ? Request.Cookies["From_date"].Value : MauIn.From_date;
            //var toDate = MauIn.To_date==""?Request.Cookies["To_Date"].Value : MauIn.To_date;
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_date"].Value;
            var MaDt = Request.Cookies["Ma_DT"] != null ? Request.Cookies["Ma_DT"].Value : string.Empty;
            DataSet ds = new DataSet();
            connectSQL();
            var SoCT = Request.Cookies["So_Ct"] != null ? Request.Cookies["So_Ct"].Value : "";
            //MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_XacNhanCKTT_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                //MauIn.From_date = Request.Cookies["From_date"].Value;
                //MauIn.To_date = Request.Cookies["To_Date"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Ma_dt", MaDt);
                    cmd.Parameters.AddWithValue("@_so_Ct", SoCT);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        
        
        public ActionResult PhieuXuatKho_Fill(MauInChungTu MauIn)
        {
            DataSet ds = new DataSet();
            connectSQL();

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_MauInChungTuSO_Detail_SAP]";

            MauIn.From_date = Request.Cookies["From_date"].Value;
            MauIn.To_date = Request.Cookies["To_Date"].Value;
            //MauIn.UserName = Request.Cookies["UserName"].Value;
            var SoCt = Request.Cookies["SoCt"].Value;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_so_Ct", SoCt);
                    cmd.Parameters.AddWithValue("@_username", MauIn.UserName);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }

        public ActionResult PhieuXuatKho_SO(MauInChungTu MauIn)
        {
            return View();
        }

        public ActionResult PhieuXuatKho_Data(MauInChungTu MauIn)
        {

            DataSet ds = new DataSet();
            connectSQL();

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_MauInChungTuSO_SAP]";
            Response.Cookies["From_date"].Value = MauIn.From_date.ToString();
            Response.Cookies["To_Date"].Value = MauIn.To_date.ToString();
            string Sales_Unit=Request.Cookies["Sales_Unit"].Value;
            MauIn.UserName = Request.Cookies["UserName"].Value;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_username", MauIn.UserName);
                    cmd.Parameters.AddWithValue("@_ma_dvcs",Sales_Unit);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult ThongBaoNoQH_Fill()
        {
            List<MauInChungTu> dmDlist = LoadDmDt("");

            ViewBag.DataItems = dmDlist;
            return View();
        }
        public ActionResult ThongBaoNoQH_In(MauInChungTu MauIn)
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            DataSet ds = new DataSet();
            connectSQL();

            //MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_ThongBaoNoQuaHan_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                //MauIn.From_date = Request.Cookies["From_date"].Value;
                //MauIn.To_date = Request.Cookies["To_Date"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {


                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_Ma_dt", MauIn.Ma_Dt);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);


                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult BangDoiChieuCongNo_Fill()

        {
            List<MauInChungTu> dmDlist = LoadDmDt("");

            ViewBag.DataItems = dmDlist;
            return View();

        }

        public List<GetData> LoadDataDoiChieuCN()
        {
            connectSQL();
            List<GetData> dataItems = new List<GetData>();
            List<GetData> dataItems2 = new List<GetData>();
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DoiChieuDoanhThuCongNo_SAP]", connection))
                {
                    var fromDate = Request.Cookies["From_date"].Value;
                    var toDate = Request.Cookies["To_Date"].Value;
                    var NgayTT = Request.Cookies["Ngay_TT"].Value;
                    var Ngay_Ky = Request.Cookies["Ngay_Ky"].Value;
                    var ma_dvcs = Request.Cookies["MA_DVCS"].Value;
                    if (ma_dvcs == "OPC_B1")
                    {
                        string ma_dvcsFirst3Chars = ma_dvcs == "OPC_B1" ? ma_dvcs.Substring(0, 3) : ma_dvcs;
                        string first3Chars = ma_dvcsFirst3Chars.Substring(0, 3);
                        ma_dvcs = first3Chars;
                    }
                    var ma_dt = Request.Cookies["MaDT"].Value;
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;
                    using (SqlDataAdapter sda = new SqlDataAdapter(command))
                    {
                        DataSet ds = new DataSet();
                        command.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                        command.Parameters.AddWithValue("@_Den_Ngay", toDate);
                        command.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                        command.Parameters.AddWithValue("@_Ma_Dt", ma_dt);
                        command.Parameters.AddWithValue("@_Ngay_TT", NgayTT);
                        command.Parameters.AddWithValue("@_Ngay_Ky", Ngay_Ky);
                        sda.Fill(ds);

                        // Kiểm tra xem DataSet có bảng dữ liệu hay không
                        if (ds.Tables.Count > 0)
                        {
                            DataTable dt = ds.Tables[1];
                            DataTable dt2 = ds.Tables[2];
                            foreach (DataRow row in dt.Rows)
                            {
                                GetData dataItem = new GetData
                                {

                                    So = row["So_Ct_hd"].ToString(),
                                    Ngay = row["Ngay_Ct_hd"].ToString(),
                                    TienHD = row["Tien_HD"].ToString(),
                                    So2 = row["So_Ct_Tt"].ToString(),
                                    Ngay1 = row["Ngay_Ct_Tt"].ToString(),
                                    SoTien = row["Tien_TT"].ToString(),
                                    CKTT = row["CKTT1"].ToString(),
                                    TongTien = row["Tong_tien"].ToString(),
                                    GhiChu1 = row["Ghi_Chu1"].ToString(),







                                };

                                dataItems.Add(dataItem);
                            }
                            //foreach (DataRow row in dt2.Rows)
                            //{
                            //    GetData dataItem2 = new GetData
                            //    {

                            //        So3 = row["So_Ct"].ToString(),
                            //        Ngay3 = row["Ngay_Ct1"].ToString(),
                            //        TienHD2 = Convert.ToDecimal(row["Ton_No1"].ToString()),
                            //        GhiChu2 = row["Ghi_Chu"].ToString(),
                            //    };

                            //    dataItems.Add(dataItem2);
                            //}
                        }
                    }
                }
            }
            return dataItems;
        }
        public ActionResult BangDoiChieuCongNo(MauInChungTu MauIn)
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            //string GopMa = Request.Cookies["GopMa"].Value;
            DataSet ds = new DataSet();
            connectSQL();
            var Ma_Dt = Request.Cookies["MaDT"].Value;
            if (ma_dvcs == "OPC_B1")
            {
                string ma_dvcsFirst3Chars = ma_dvcs == "OPC_B1" ? ma_dvcs.Substring(0, 3) : ma_dvcs;
                string first3Chars = ma_dvcsFirst3Chars.Substring(0, 3);
                ma_dvcs = first3Chars;
            }
            //MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_DoiChieuDoanhThuCongNo_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                //MauIn.From_date = Request.Cookies["From_date"].Value;
                //MauIn.To_date = Request.Cookies["To_Date"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {


                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_Ma_dt", Ma_Dt);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                    cmd.Parameters.AddWithValue("@_Ngay_TT", MauIn.Ngay_TT);
                    cmd.Parameters.AddWithValue("@_Ngay_Ky", MauIn.Ngay_Ky);
                    cmd.Parameters.AddWithValue("@_So", MauIn.So);


                    sda.Fill(ds);

                }
            }



            return View(ds);
        }


        public List<GetData> LoadDataTBNoQH()
        {
            connectSQL();
            List<GetData> dataItems = new List<GetData>();
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_ThongBaoNoQuaHan_SAP]", connection))
                {
                    var fromDate = Request.Cookies["From_date"].Value;
                    var toDate = Request.Cookies["To_Date"].Value;

                    var ma_dvcs = Request.Cookies["MA_DVCS"].Value;
                    var ma_dt = Request.Cookies["Ma_Dt"].Value;
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;
                    using (SqlDataAdapter sda = new SqlDataAdapter(command))
                    {
                        DataSet ds = new DataSet();
                        command.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                        command.Parameters.AddWithValue("@_Den_Ngay", toDate);
                        command.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                        command.Parameters.AddWithValue("@_Ma_Dt", ma_dt);
                        sda.Fill(ds);

                        // Kiểm tra xem DataSet có bảng dữ liệu hay không
                        if (ds.Tables.Count > 0)
                        {
                            DataTable dt = ds.Tables[1];

                            foreach (DataRow row in dt.Rows)
                            {
                                GetData dataItem = new GetData
                                {
                                    SoHD = row["So_Ct"].ToString(),
                                    NgayXuat = row["Ngay_Ct1"].ToString(),
                                    HanTT = row["Han_Thanh_Toan"].ToString(),

                                    NgayQH = Convert.ToInt32(row["So_Ngay_Qua_Han"].ToString()),
                                    TienNo = Convert.ToDecimal(row["Tong_No"].ToString()),





                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }
            return dataItems;
        }
        public ActionResult ExportToExcel()
        {
            var fileName = $"MauThongBaoNoQH{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            // Lấy dữ liệu từ cookie
            List<GetData> combinedData = LoadDataTBNoQH();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("MySheet");
                worksheet.View.ShowGridLines = false;

                // ... (Các bước tạo nội dung tệp Excel như bạn đã làm)
                // Đường dẫn đến hình ảnh trong thư mục 'image'
                var imagePath = Server.MapPath("~/assets/images/logo.png"); // Thay thế bằng đường dẫn thật
                                                                            // Lấy giá trị từ biến Dvcs
                string Dvcs = Request.Cookies["Dvcs"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs"].Value) : "";
                string Dvcs1 = Request.Cookies["Dvcs1"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs1"].Value) : "";
                string ten_dt = Request.Cookies["ten_dt"] != null ? HttpUtility.UrlDecode(Request.Cookies["ten_dt"].Value) : "";
                string denngay = Request.Cookies["DenNgayCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["DenNgayCookie"].Value) : "";
                string tongno = Request.Cookies["TongNo"] != null ? HttpUtility.UrlDecode(Request.Cookies["TongNo"].Value) : "";
                string QuaHan = Request.Cookies["QuaHan"] != null ? HttpUtility.UrlDecode(Request.Cookies["QuaHan"].Value) : "";
                string HanNgay = Request.Cookies["HanNgay"] != null ? HttpUtility.UrlDecode(Request.Cookies["HanNgay"].Value) : "";
                string CN = Request.Cookies["CN"] != null ? HttpUtility.UrlDecode(Request.Cookies["CN"].Value) : "";
                string TK = Request.Cookies["TK"] != null ? HttpUtility.UrlDecode(Request.Cookies["TK"].Value) : "";
                string LH = Request.Cookies["LH"] != null ? HttpUtility.UrlDecode(Request.Cookies["LH"].Value) : "";
                // Đặt font chữ "Arial" cho toàn bộ tệp Excel
                worksheet.Cells.Style.Font.Name = "Times New Roman";

                // Chèn hình ảnh từ tệp hình vào ô A1
                ExcelPicture picture = worksheet.Drawings.AddPicture("MyPicture", new FileInfo(imagePath));
                picture.SetSize(55, 45); // Đặt kích thước cho hình ảnh
                picture.From.Row = 1;
                picture.From.Column = 0;
                worksheet.Column(1).Width = 8;

                // Đặt văn bản vào ô A2
                worksheet.Cells["B1"].Value = "CTY CỔ PHẦN DƯỢC PHẨM OPC";
                var cellB1 = worksheet.Cells["B1"];
                cellB1.Style.Font.Bold = true;
                worksheet.Cells["B1"].Style.Indent = 3;
                worksheet.Cells["B2"].Style.Indent = 3;
                worksheet.Cells["B3"].Style.Indent = 3;
                worksheet.Cells["B2"].Value = Dvcs;
                worksheet.Cells["B3"].Value = $"Số:............................/KT-{Dvcs1}";
                worksheet.Cells["H1"].Value = "Cộng Hòa Xã Hội Chủ Nghĩa Việt Nam";
                worksheet.Cells["H2"].Value = "Độc Lập - Tự Do - Hạnh Phúc";
                worksheet.Cells["H2"].Style.Indent = 4;
                worksheet.Cells["H2"].Style.Font.UnderLine = true;
                worksheet.Cells["E4"].Value = "THÔNG BÁO NỢ QUÁ HẠN";
                worksheet.Cells["E4"].Style.Font.Bold = true;
                worksheet.Cells["E4"].Style.Font.Size = 16;
                worksheet.Cells["A6"].Value = $"Kính gửi: {ten_dt}";
                worksheet.Cells["A6"].Style.Font.Bold = true;
                worksheet.Cells["A8"].Value = $"{Dvcs} - Công ty Cổ Phần Dược Phẩm OPC trân trọng thông báo đến quý khách hàng có số dư nợ mà Quý Khách";
                worksheet.Cells["A9"].Value = $"hàng chưa thanh toán cho chúng tôi tính đến ngày {denngay} là: {tongno}";
                worksheet.Cells["B11"].Value = $"Trong đó nợ quá hạn là: {QuaHan} bao gồm các hóa đơn sau:";
                var startRow = 13;
                var startColumn = 1;
                worksheet.Cells[startRow - 1, startColumn].Value = "STT";
                worksheet.Cells[startRow - 1, startColumn + 1].Value = "SỐ HÓA ĐƠN";
                worksheet.Cells[startRow - 1, startColumn + 2].Value = "NGÀY XUẤT";
                worksheet.Cells[startRow - 1, startColumn + 3].Value = "TIỀN NỢ";
                worksheet.Cells[startRow - 1, startColumn + 4].Value = "HẠN THANH TOÁN";
                worksheet.Cells[startRow - 1, startColumn + 5].Value = "NGÀY QUÁ HẠN";
                for (int col = 0; col < 6; col++)
                {
                    var columnHeaderCell = worksheet.Cells[startRow - 1, startColumn + col];
                    columnHeaderCell.Style.Font.Bold = true;
                    columnHeaderCell.Style.Font.Size = 10;
                    columnHeaderCell.Style.Font.Color.SetColor(Color.Black);
                    columnHeaderCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    columnHeaderCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    columnHeaderCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    columnHeaderCell.Style.Fill.BackgroundColor.SetColor(Color.White);
                }
                var columnHeaderStyle = worksheet.Cells[startRow - 1, startColumn, startRow - 1, startColumn + 5].Style;
                columnHeaderStyle.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black); // Đóng khung solid đen
                worksheet.Column(startColumn).Width = 5; // Độ rộng cột cho "STT"
                worksheet.Column(startColumn + 1).Width = 15; // Độ rộng cột cho "SỐ HÓA ĐƠN"
                worksheet.Column(startColumn + 2).Width = 15; // Độ rộng cột cho "NGÀY XUẤT"
                worksheet.Column(startColumn + 3).Width = 15; // Độ rộng cột cho "TIỀN NỢ"
                worksheet.Column(startColumn + 4).Width = 18; // Độ rộng cột cho "HẠN THANH TOÁN"
                worksheet.Column(startColumn + 5).Width = 15; // 

                // Đảm bảo rằng có dữ liệu trong biến tableData
                if (combinedData != null && combinedData.Any())
                {
                    var stt = 1;
                    // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
                    for (int row = 0; row < combinedData.Count; row++)
                    {
                        var rowData = combinedData[row];

                        worksheet.Cells[startRow + row, startColumn].Value = stt;

                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn]);
                        worksheet.Cells[startRow + row, startColumn + 1].Value = rowData.SoHD;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 1]);
                        worksheet.Cells[startRow + row, startColumn + 2].Value = rowData.NgayXuat;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 2]);
                        worksheet.Cells[startRow + row, startColumn + 3].Value = rowData.TienNo;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 3]);
                        worksheet.Cells[startRow + row, startColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[startRow + row, startColumn + 4].Value = rowData.HanTT;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 4]);
                        worksheet.Cells[startRow + row, startColumn + 5].Value = rowData.NgayQH;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 5]);
                        stt++;
                    }

                }
                else
                {
                    worksheet.Cells[startRow, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
                }
                worksheet.Cells[startRow + combinedData.Count, startColumn + 1].Value = "Tổng cộng";
                worksheet.Cells[startRow + combinedData.Count, startColumn + 1].Style.Font.Bold = true;
                worksheet.Cells[startRow + combinedData.Count, startColumn + 3].Value = $"{QuaHan}"; // Ví dụ: Ghi giá trị tổng vào cột thứ 4
                worksheet.Cells[startRow + combinedData.Count, startColumn + 3].Style.Font.Bold = true;
                worksheet.Cells[startRow + combinedData.Count, startColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[startRow + combinedData.Count, startColumn].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count, startColumn + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count, startColumn + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count, startColumn + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count, startColumn + 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count, startColumn + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                // Xóa hàng tiêu đề mặc định

                var dataRowStyle = worksheet.Cells[startRow, startColumn, startRow, startColumn + 5].Style;
                dataRowStyle.Font.Bold = false;
                dataRowStyle.Font.Color.SetColor(Color.Black);
                dataRowStyle.Fill.PatternType = ExcelFillStyle.None;
                // Tạo bảng trong tệp Excel
                var endRow = startRow + combinedData.Count;


                //var tableRange = worksheet.Cells[startRow, startColumn, endRow, endColumn];
                //var table = worksheet.Tables.Add(tableRange, "MyTable");
                //table.TableStyle = TableStyles.Light1;
                int nextRow = endRow + 1;
                worksheet.Cells[nextRow, startColumn].Value = $"Kính đề nghị Quý khách vui lòng đối chiếu và xác nhận số tiền gửi về {Dvcs} - Công Ty Cổ Phần Dược Phẩm OPC";
                worksheet.Cells[nextRow + 1, startColumn].Value = $"trước ngày {HanNgay}.Đồng thời sớm thanh toán số dư nợ quá hạn trên cho Chi Nhánh chúng tôi bằng tiền mặt hoặc chuyển vào";
                worksheet.Cells[nextRow + 2, startColumn].Value = $" tài khoản: Chi nhánh Công Ty Cổ Phẩn Dược Phẩm OPC tại {CN}.";
                worksheet.Cells[nextRow + 3, startColumn].Value = $"Số tài khoản: {TK}";
                worksheet.Cells[nextRow + 3, startColumn].Style.Indent = 2;
                worksheet.Cells[nextRow + 4, startColumn].Value = $"Khi cần đối chiếu xin liên hệ {LH}";
                worksheet.Cells[nextRow + 4, startColumn].Style.Indent = 2;
                worksheet.Cells[nextRow + 6, startColumn].Value = "Trân trọng!";
                worksheet.Cells[nextRow + 6, startColumn].Style.Indent = 2;
                worksheet.Cells[nextRow + 6, startColumn].Style.Font.Italic = true;
                worksheet.Cells[nextRow + 8, startColumn + 1].Value = "Khách Hàng Xác Nhận";
                worksheet.Cells[nextRow + 8, startColumn + 1].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 8, startColumn + 4].Value = "Giám Đốc";
                worksheet.Cells[nextRow + 8, startColumn + 4].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 8, startColumn + 7].Value = "Kế Toán";
                worksheet.Cells[nextRow + 8, startColumn + 7].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 9, startColumn].Value = "(Ký, đóng dấu, ghi rõ họ tên)";
                worksheet.Cells[nextRow + 9, startColumn].Style.Indent = 4;
                worksheet.Cells[nextRow + 9, startColumn].Style.Font.Italic = true;

                package.Save();
                byte[] fileBytes = package.GetAsByteArray();

                // Trả về tệp Excel dưới dạng dữ liệu binary
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

            }



            return View("ThongBaoNoQH_In");
        }

        public ActionResult ExportBaoCaoCongNo()
        {



            var fileName = $"BangDoiChieuCongNo{DateTime.Now:yyyyMMddHHmmss}.xlsx";

            // Kiểm tra xem có dữ liệu từ cookie không
            List<GetData> combinedData = LoadDataDoiChieuCN();
            List<GetData> combinedDataHD = LoadDataDTCN2();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("MySheet");
                worksheet.View.ShowGridLines = false;
                var startRow = 13;
                var startColumn = 1;
                // ... (Các bước tạo nội dung tệp Excel như bạn đã làm)
                // Đường dẫn đến hình ảnh trong thư mục 'image'
                var imagePath = Server.MapPath("~/assets/images/logo.png"); // Thay thế bằng đường dẫn thật
                                                                            // Lấy giá trị từ biến Dvcs
                string Dvcs = Request.Cookies["Dvcs"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs"].Value) : "";
                string TuNgay = Request.Cookies["tungay"] != null ? HttpUtility.UrlDecode(Request.Cookies["tungay"].Value) : "";
                string TuThang = Request.Cookies["tuthang"] != null ? HttpUtility.UrlDecode(Request.Cookies["tuthang"].Value) : "";
                string DenNgay = Request.Cookies["denngay"] != null ? HttpUtility.UrlDecode(Request.Cookies["denngay"].Value) : "";
                string DenThang = Request.Cookies["denthang"] != null ? HttpUtility.UrlDecode(Request.Cookies["denthang"].Value) : "";
                string Nam = Request.Cookies["nam"] != null ? HttpUtility.UrlDecode(Request.Cookies["nam"].Value) : "";
                string DiaChi = Request.Cookies["Dia_Chi"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dia_Chi"].Value) : "";
                string NoDauKy = Request.Cookies["NoDauKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["NoDauKy"].Value) : "";
                string TienHD = Request.Cookies["TienHD"] != null ? HttpUtility.UrlDecode(Request.Cookies["TienHD"].Value) : "";
                string TonNo = Request.Cookies["TonNo"] != null ? HttpUtility.UrlDecode(Request.Cookies["TonNo"].Value) : "";
                string TienChu = Request.Cookies["TienChu"] != null ? HttpUtility.UrlDecode(Request.Cookies["TienChu"].Value) : "";
                string TonNo2 = Request.Cookies["TonNo2"] != null ? HttpUtility.UrlDecode(Request.Cookies["TonNo2"].Value) : "";
                string NgayTT = Request.Cookies["NgayTT"] != null ? HttpUtility.UrlDecode(Request.Cookies["NgayTT"].Value) : "";
                string ChiNhanh = Request.Cookies["ChiNhanh"] != null ? HttpUtility.UrlDecode(Request.Cookies["ChiNhanh"].Value) : "";
                string DiaChi2 = Request.Cookies["DiaChi2"] != null ? HttpUtility.UrlDecode(Request.Cookies["DiaChi2"].Value) : "";
                //string Dvcs1 = Request.Cookies["Dvcs1"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs1"].Value) : "";
                string ten_kh = Request.Cookies["Ten_Dt"] != null ? HttpUtility.UrlDecode(Request.Cookies["Ten_Dt"].Value) : "";
                string NgayKy = Request.Cookies["NgayKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["NgayKy"].Value) : "";
                string ThangKy = Request.Cookies["ThangKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["ThangKy"].Value) : "";
                string NamKy = Request.Cookies["NamKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["NamKy"].Value) : "";
                //string HanNgay = Request.Cookies["HanNgay"] != null ? HttpUtility.UrlDecode(Request.Cookies["HanNgay"].Value) : "";
                string CN = Request.Cookies["CN"] != null ? HttpUtility.UrlDecode(Request.Cookies["CN"].Value) : "";
                string TK = Request.Cookies["TK"] != null ? HttpUtility.UrlDecode(Request.Cookies["TK"].Value) : "";
                string LH = Request.Cookies["LH"] != null ? HttpUtility.UrlDecode(Request.Cookies["LH"].Value) : "";
                string Time = Request.Cookies["Time"] != null ? HttpUtility.UrlDecode(Request.Cookies["Time"].Value) : "";

                string TienTT = Request.Cookies["Tien_TT"] != null ? HttpUtility.UrlDecode(Request.Cookies["Tien_TT"].Value) : "";
                string CKTT = Request.Cookies["CKTT"] != null ? HttpUtility.UrlDecode(Request.Cookies["CKTT"].Value) : "";
                string TCCK = Request.Cookies["TC_CKTT_TienTT"] != null ? HttpUtility.UrlDecode(Request.Cookies["TC_CKTT_TienTT"].Value) : "";
                // Đặt font chữ "Arial" cho toàn bộ tệp Excel
                worksheet.Cells.Style.Font.Name = "Times New Roman";

                // Chèn hình ảnh từ tệp hình vào ô A1
                ExcelPicture picture = worksheet.Drawings.AddPicture("MyPicture", new FileInfo(imagePath));
                picture.SetSize(70, 50); // Đặt kích thước cho hình ảnh
                picture.From.Row = 1;
                picture.From.Column = 0;

                worksheet.Column(1).Width = 8;

                // Đặt văn bản vào ô A2
                worksheet.Cells["B1"].Value = "CTY CỔ PHẦN DƯỢC PHẨM OPC";
                var cellB1 = worksheet.Cells["B1"];
                cellB1.Style.Font.Bold = true;
                //worksheet.Cells["B1"].Style.Indent = 3;
                //worksheet.Cells["B2"].Style.Indent = 3;
                //worksheet.Cells["B3"].Style.Indent = 3;
                worksheet.Cells["B2"].Value = Dvcs;
                worksheet.Cells["B3"].Value = $"Số:";
                worksheet.Cells["H1"].Value = "Cộng Hòa Xã Hội Chủ Nghĩa Việt Nam";
                worksheet.Cells["H2"].Value = "Độc Lập - Tự Do - Hạnh Phúc";
                worksheet.Cells["H2"].Style.Indent = 4;

                worksheet.Cells["E4"].Value = "BẢNG ĐỐI CHIẾU DOANH THU CÔNG NỢ";
                worksheet.Cells["E4"].Style.Font.Bold = true;
                worksheet.Cells["E4"].Style.Font.Size = 16;
                worksheet.Cells["E5"].Value = $"{Time}";
                worksheet.Cells["E5"].Style.Indent = 4;
                worksheet.Cells["A7"].Value = $"Tên khách hàng: {ten_kh}";
                worksheet.Cells["A7"].Style.Font.Bold = true;
                worksheet.Cells["A8"].Value = $"Địa chỉ khách hàng: {DiaChi}";
                worksheet.Cells["A8"].Style.Font.Bold = true;
                worksheet.Cells["A9"].Value = $"I.Số dư nợ trước ngày: {TuNgay}/{TuThang}/{Nam}";
                worksheet.Cells["A9"].Style.Font.Bold = true;
                worksheet.Cells["E9"].Value = $"mang sang {NoDauKy} đồng";
                worksheet.Cells["E9"].Style.Font.Bold = true;
                worksheet.Cells["A10"].Value = "II.Doanh thu và công nợ phát sinh trong kỳ đối chiếu này: ";
                worksheet.Cells["A10"].Style.Font.Bold = true;
                //worksheet.Cells["A8"].Value = $"{Dvcs} - Công ty Cổ Phần Dược Phẩm OPC trân trọng thông báo đến quý khách hàng có số dư nợ mà Quý Khách";
                //worksheet.Cells["A9"].Value = $"hàng chưa thanh toán cho chúng tôi tính đến ngày {denngay} là: {tongno}";
                //worksheet.Cells["B11"].Value = $"Trong đó nợ quá hạn là: {QuaHan} bao gồm các hóa đơn sau:";

                var sttCell = worksheet.Cells[startRow - 1, startColumn, startRow, startColumn];
                sttCell.Merge = true;
                sttCell.Value = "STT";
                sttCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sttCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
                sttCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc
                sttCell.Style.Font.Bold = true;
                sttCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                var KHMua = worksheet.Cells[startRow - 1, startColumn + 1, startRow - 1, startColumn + 3];
                KHMua.Merge = true;
                KHMua.Value = "KHÁCH HÀNG MUA";
                KHMua.Style.Font.Bold = true;
                KHMua.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                KHMua.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                KHMua.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow, startColumn + 1].Value = "SỐ";
                worksheet.Cells[startRow, startColumn + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow, startColumn + 1].Style.Font.Bold = true;
                worksheet.Cells[startRow, startColumn + 2].Value = "NGÀY";
                worksheet.Cells[startRow, startColumn + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow, startColumn + 2].Style.Font.Bold = true;
                worksheet.Cells[startRow, startColumn + 3].Value = "SỐ TIỀN";
                worksheet.Cells[startRow, startColumn + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow, startColumn + 3].Style.Font.Bold = true;
                var KHTT = worksheet.Cells[startRow - 1, startColumn + 4, startRow - 1, startColumn + 8];
                KHTT.Merge = true;
                KHTT.Value = "KHÁCH HÀNG THANH TOÁN/TRẢ HÀNG BÙ TRỪ";
                KHTT.Style.Font.Bold = true;
                KHTT.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                //worksheet.Column(startColumn + 4).Width = 20;
                worksheet.Cells[startRow - 1, startColumn + 9].Value = "";
                worksheet.Cells[startRow - 1, startColumn + 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow, startColumn + 4].Value = "SỐ";
                worksheet.Cells[startRow, startColumn + 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow, startColumn + 4].Style.Font.Bold = true;

                worksheet.Cells[startRow, startColumn + 5].Value = "NGÀY";
                worksheet.Cells[startRow, startColumn + 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow, startColumn + 5].Style.Font.Bold = true;
                worksheet.Cells[startRow, startColumn + 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells[startRow, startColumn + 6].Value = "SỐ TIỀN";
                worksheet.Cells[startRow, startColumn + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow, startColumn + 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 6].Style.Font.Bold = true;

                worksheet.Cells[startRow, startColumn + 7].Value = "CKTT";
                worksheet.Cells[startRow, startColumn + 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow, startColumn + 7].Style.Font.Bold = true;

                worksheet.Cells[startRow, startColumn + 8].Value = "TỔNG TIỀN";
                worksheet.Cells[startRow, startColumn + 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow, startColumn + 8].Style.Font.Bold = true;
                worksheet.Cells[startRow, startColumn + 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells[startRow, startColumn + 9].Value = "GHI CHÚ";
                worksheet.Cells[startRow, startColumn + 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[startRow, startColumn + 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow, startColumn + 9].Style.Font.Bold = true;
                if (combinedData != null && combinedData.Any())
                {
                    var stt = 1;
                    // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
                    for (int row = 0; row < combinedData.Count; row++)
                    {
                        var rowData = combinedData[row];
                        FormatCellNoQH(worksheet.Cells[startRow + row + 1, startColumn]);
                        worksheet.Cells[startRow + row + 1, startColumn].Value = stt;
                        worksheet.Cells[startRow + row + 1, startColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[startRow + row + 1, startColumn + 1].Value = rowData.So;
                        FormatCellNoQH(worksheet.Cells[startRow + row + 1, startColumn + 1]);
                        worksheet.Cells[startRow + row + 1, startColumn + 2].Value = rowData.Ngay;
                        FormatCellNoQH(worksheet.Cells[startRow + row + 1, startColumn + 2]);
                        worksheet.Cells[startRow + row + 1, startColumn + 3].Value = rowData.TienHD;
                        worksheet.Cells[startRow + row + 1, startColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        FormatCellNoQH(worksheet.Cells[startRow + row + 1, startColumn + 3]);

                        worksheet.Cells[startRow + row + 1, startColumn + 4].Value = rowData.So2;
                        FormatCellNoQH(worksheet.Cells[startRow + row + 1, startColumn + 4]);

                        worksheet.Cells[startRow + row + 1, startColumn + 5].Value = rowData.Ngay1;
                        FormatCellNoQH(worksheet.Cells[startRow + row + 1, startColumn + 5]);
                        worksheet.Cells[startRow + row + 1, startColumn + 6].Value = rowData.SoTien;
                        FormatCellNoQH(worksheet.Cells[startRow + row + 1, startColumn + 6]);
                        worksheet.Cells[startRow + row + 1, startColumn + 7].Value = rowData.CKTT;
                        FormatCellNoQH(worksheet.Cells[startRow + row + 1, startColumn + 7]);
                        worksheet.Cells[startRow + row + 1, startColumn + 8].Value = rowData.TongTien;
                        FormatCellNoQH(worksheet.Cells[startRow + row + 1, startColumn + 8]);
                        worksheet.Cells[startRow + row + 1, startColumn + 9].Value = rowData.GhiChu1;
                        FormatCellNoQH(worksheet.Cells[startRow + row + 1, startColumn + 9]);
                        stt++;
                    }
                }
                else
                {
                    worksheet.Cells[startRow, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
                }
                var TC = worksheet.Cells[startRow + combinedData.Count + 1, startColumn, startRow + combinedData.Count + 1, startColumn + 2];
                TC.Merge = true;
                TC.Value = "Tổng cộng: ";
                TC.Style.Font.Bold = true;
                TC.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                TC.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
                TC.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 3].Value = $"{TienHD}"; // Ví dụ: Ghi giá trị tổng vào cột thứ 4
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 3].Style.Font.Bold = true;
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 6].Value = $"{TienTT}";
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 7].Value = $"{CKTT}";
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 8].Value = $"{TCCK}";
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 6].Style.Font.Bold = true;
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 7].Style.Font.Bold = true;
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 8].Style.Font.Bold = true;
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count + 1, startColumn + 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                var endII = startRow + combinedData.Count + 1;
                var nextII = endII + 1;
                worksheet.Cells[nextII, startColumn].Value = $"III. Số tiền khách hàng chưa thanh toán, tính đến cuối ngày: {DenNgay}/{DenThang}/{Nam} là: {TonNo} đồng";
                worksheet.Cells[nextII, startColumn].Style.Font.Bold = true;
                worksheet.Cells[nextII + 1, startColumn].Value = $"Số tiền bằng chữ: {TienChu}";
                worksheet.Cells[nextII + 1, startColumn].Style.Font.Bold = true;
                worksheet.Cells[nextII + 2, startColumn].Value = "Chi tiết các hóa đơn chưa thanh toán: ";

                var startRowIII = nextII + 3;

                var sttIIICell = worksheet.Cells[startRowIII + 1, startColumn, startRowIII + 2, startColumn + 1];
                sttIIICell.Merge = true;
                worksheet.Column(startColumn).Width = 15; // Đặt chiều rộng của cột chứa ô "STT" thành 15 đơn vị.

                sttIIICell.Value = "STT";
                sttIIICell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
                sttIIICell.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc
                sttIIICell.Style.Font.Bold = true;
                sttIIICell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                var HD = worksheet.Cells[startRowIII + 1, startColumn + 2, startRowIII + 1, startColumn + 8];
                HD.Merge = true;
                HD.Value = "HÓA ĐƠN";
                HD.Style.Font.Bold = true;
                HD.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                HD.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                HD.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                var SO = worksheet.Cells[startRowIII + 2, startColumn + 2, startRowIII + 2, startColumn + 3];
                SO.Merge = true;
                SO.Value = "SỐ";
                worksheet.Column(startColumn + 1).Width = 15;
                SO.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                SO.Style.Font.Bold = true;
                SO.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                SO.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                var NGAY = worksheet.Cells[startRowIII + 2, startColumn + 4, startRowIII + 2, startColumn + 5];
                NGAY.Merge = true;
                NGAY.Value = "NGÀY";
                worksheet.Column(startColumn + 2).Width = 15;
                NGAY.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                NGAY.Style.Font.Bold = true;
                NGAY.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                NGAY.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                var TIENHD = worksheet.Cells[startRowIII + 2, startColumn + 6];
                //TIENHD.Merge = true;
                TIENHD.Value = "SỐ TIỀN HD";
                worksheet.Column(startColumn + 3).Width = 15;
                TIENHD.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                TIENHD.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TIENHD.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                TIENHD.Style.Font.Bold = true;

                var GHICHU = worksheet.Cells[startRowIII + 2, startColumn + 7, startRowIII + 2, startColumn + 8];
                GHICHU.Merge = true;
                GHICHU.Value = "GHI CHÚ";
                worksheet.Column(startColumn + 4).Width = 15;
                worksheet.Column(startColumn + 5).Width = 15;
                worksheet.Column(startColumn + 6).Width = 15;
                worksheet.Column(startColumn + 8).Width = 15;
                GHICHU.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                GHICHU.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                GHICHU.Style.Font.Bold = true;
                GHICHU.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //// Bỏ gộp ô trước khi xóa
                //worksheet.Cells["A21:B22"].Merge = false;

                //// Xóa ô hoặc thực hiện các thao tác khác
                //worksheet.Cells["A21:B22"].Clear();

                if (combinedDataHD != null && combinedDataHD.Any())
                {
                    // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
                    var stt = 1;
                    for (int row = 0; row < combinedDataHD.Count; row++)
                    {
                        var rowData = combinedDataHD[row];
                        //worksheet.Cells[startRowIII + row+3, startColumn].Value = stt;
                        var sttCell2 = worksheet.Cells[startRowIII + row + 3, startColumn, startRowIII + row + 3, startColumn + 1];
                        sttCell2.Merge = true;
                        sttCell2.Value = stt;
                        FormatCellNoQH(sttCell2);


                        var soCell = worksheet.Cells[startRowIII + row + 3, startColumn + 2, startRowIII + row + 3, startColumn + 3];
                        soCell.Merge = true;
                        soCell.Value = rowData.So;
                        FormatCellNoQH(soCell);
                        //worksheet.Cells[startRowIII + row + 3, startColumn+1].Value = rowData.So;
                        //FormatCellNoQH(worksheet.Cells[startRowIII + row + 3, startColumn+1]);

                        var ngayCell = worksheet.Cells[startRowIII + row + 3, startColumn + 4, startRowIII + row + 3, startColumn + 5];
                        ngayCell.Merge = true;
                        ngayCell.Value = rowData.Ngay;
                        FormatCellNoQH(ngayCell);

                        var tienhdCell = worksheet.Cells[startRowIII + row + 3, startColumn + 6];
                        //tienhdCell.Merge = true;
                        tienhdCell.Value = rowData.TienHD;
                        worksheet.Column(startColumn + 6).Width = 30;
                        FormatCellNoQH(tienhdCell);
                        tienhdCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        //worksheet.Cells[startRowIII + row + 3, startColumn + 2].Value = rowData.Ngay;
                        //FormatCellNoQH(worksheet.Cells[startRowIII + row + 3, startColumn+2]);
                        //worksheet.Cells[startRowIII + row + 3, startColumn + 3].Value = rowData.TienHD;
                        //FormatCellNoQH(worksheet.Cells[startRowIII + row + 3, startColumn+3]);

                        var ghichu = worksheet.Cells[startRowIII + row + 3, startColumn + 7, startRowIII + row + 3, startColumn + 8];
                        ghichu.Merge = true;
                        ghichu.Value = rowData.GhiChu;
                        FormatCellNoQH(ghichu);
                        //worksheet.Cells[startRowIII + row + 3, startColumn + 4].Value = rowData.GhiChu;
                        //FormatCellNoQH(worksheet.Cells[startRowIII + row + 3, startColumn + 4]);
                        //var sttCell = worksheet.Cells[startRow + 3 + row, startColumn, startRow + 3 + row, startColumn + 1];
                        //sttCell.Merge = true;
                        //sttCell.Value = stt;
                        //FormatCellNoQH(sttCell);

                        //var soCell = worksheet.Cells[startRow + 3 + row, startColumn + 2, startRow + 3 + row, startColumn + 3];
                        //soCell.Merge = true;
                        //soCell.Value = rowData.So;
                        //FormatCellNoQH(soCell);

                        //var ngayCell = worksheet.Cells[startRow + 3 + row, startColumn + 4, startRow + 3 + row, startColumn + 5];
                        //ngayCell.Merge = true;
                        //ngayCell.Value = rowData.Ngay;
                        //FormatCellNoQH(ngayCell);

                        //var tienhdCell = worksheet.Cells[startRow + 3 + row, startColumn + 6];
                        ////tienhdCell.Merge = true;
                        //tienhdCell.Value = rowData.TienHD;
                        //worksheet.Column(startColumn + 6).Width = 30;
                        //FormatCellNoQH(tienhdCell);
                        //tienhdCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        //var codekm = worksheet.Cells[startRow + 3 + row, startColumn + 7, startRow + 3 + row, startColumn + 8];
                        //codekm.Merge = true;

                        //FormatCellNoQH(codekm);

                        //var ghichu = worksheet.Cells[startRow + 3 + row, startColumn + 9, startRow + 3 + row, startColumn + 10];
                        //ghichu.Merge = true;
                        //ghichu.Value = rowData.GhiChu;
                        //FormatCellNoQH(ghichu);
                        stt++;
                    }
                }
                else
                {
                    worksheet.Cells[startRowIII, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
                }
                var TC2 = worksheet.Cells[startRowIII + 2 + combinedDataHD.Count + 1, startColumn, startRowIII + 2 + combinedDataHD.Count + 1, startColumn + 5];
                TC2.Merge = true;
                TC2.Value = "Tổng cộng: ";
                TC2.Style.Font.Bold = true;
                TC2.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                TC2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
                TC2.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc

                var TONNO2 = worksheet.Cells[startRowIII + 2 + combinedDataHD.Count + 1, startColumn + 6];
                //TONNO2.Merge = true;
                TONNO2.Value = $"{TonNo2}";
                //worksheet.Cells[startRowIII + 1 + combinedDataHD.Count, startColumn + 6].Value = $"{TonNo2}";
                TONNO2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                TONNO2.Style.Font.Bold = true;
                TONNO2.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRowIII + 2 + combinedDataHD.Count, startColumn + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                var lastCell = worksheet.Cells[startRowIII + 2 + combinedDataHD.Count + 1, startColumn + 7, startRowIII + 2 + combinedDataHD.Count + 1, startColumn + 8];
                lastCell.Merge = true;
                lastCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                var endRowIII = startRowIII + 2 + combinedDataHD.Count + 2;
                worksheet.Cells[endRowIII, startColumn].Value = $"Xin vui lòng xác nhận và gửi lại cho {Dvcs} trước ngày {NgayTT}";
                worksheet.Cells[endRowIII, startColumn].Style.Font.Bold = true;
                worksheet.Cells[endRowIII + 1, startColumn].Value = $"Nơi nhận: {ChiNhanh}";
                worksheet.Cells[endRowIII + 1, startColumn].Style.Font.Bold = true;
                worksheet.Cells[endRowIII + 2, startColumn].Value = $"Địa chỉ: {DiaChi2}";
                worksheet.Cells[endRowIII + 2, startColumn].Style.Font.Bold = true;
                worksheet.Cells[endRowIII + 3, startColumn].Value = $"Khi cần đối chiếu số liệu liên hệ: {LH}";
                worksheet.Cells[endRowIII + 3, startColumn].Style.Font.Bold = true;
                worksheet.Cells[endRowIII + 4, startColumn].Value = $"Số tiền còn nợ đề nghị Quý khách hàng thanh toán bằng tiền mặt hoặc chuyển khoản vào tài khoản {CN}, số";
                worksheet.Cells[endRowIII + 4, startColumn].Style.Font.Bold = true;
                worksheet.Cells[endRowIII + 5, startColumn].Value = $"tài khoản: {TK}";
                worksheet.Cells[endRowIII + 5, startColumn].Style.Font.Bold = true;
                worksheet.Cells[endRowIII + 7, startColumn].Value = "Trân trọng cảm ơn!";
                worksheet.Cells[endRowIII + 7, startColumn].Style.Font.Bold = true;
                worksheet.Cells[endRowIII + 7, startColumn].Style.Font.Italic = true;
                worksheet.Cells[endRowIII + 8, startColumn + 7].Value = $"Ngày {NgayKy} tháng {ThangKy} năm {NamKy}";
                worksheet.Cells[endRowIII + 8, startColumn + 7].Style.Font.Bold = true;
                worksheet.Cells[endRowIII + 9, startColumn].Value = "ĐẠI DIỆN KHÁCH HÀNG";
                worksheet.Cells[endRowIII + 9, startColumn].Style.Font.Bold = true;
                worksheet.Cells[endRowIII + 9, startColumn + 7].Value = "ĐẠI DIỆN CHI NHÁNH";
                worksheet.Cells[endRowIII + 9, startColumn + 7].Style.Font.Bold = true;
                worksheet.Cells[endRowIII + 9, startColumn + 7].Style.Indent = 2;



                package.Save();
                byte[] fileBytes = package.GetAsByteArray();

                // Trả về tệp Excel dưới dạng dữ liệu binary
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

            }





        }




        //Ham chay duoc 


        //public ActionResult ExportBaoCaoCongNo()
        //{

        //    var fileName = $"BangDoiChieuCongNo{DateTime.Now:yyyyMMddHHmmss}.xlsx";
        //    // Lấy dữ liệu từ cookie
        //    //string jsonData = Request.Cookies["tableDataCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookie"].Value) : "";
        //    //string jsonData2 = Request.Cookies["tableData2Cookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableData2Cookie"].Value) : "";

        //    string cookie1Value = Request.Cookies["tableDataCookie1"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookie1"].Value) : "";
        //    string cookie2Value = Request.Cookies["tableDataCookie2"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookie2"].Value) : "";
        //    string cookie3Value = Request.Cookies["tableDataCookie3"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookie3"].Value) : "";
        //    //string cookie4Value = Request.Cookies["tableDataCookie4"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookie4"].Value) : "";
        //    //string cookie5Value = Request.Cookies["tableDataCookie5"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookie5"].Value) : "";

        //    string cookieHD1Value = Request.Cookies["tableDataCookieHD1"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookieHD1"].Value) : "";
        //    string cookieHD2Value = Request.Cookies["tableDataCookieHD2"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookieHD2"].Value) : "";
        //    //string cookieHD3Value = Request.Cookies["tableDataCookieHD3"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookieHD3"].Value) : "";
        //    //string cookieHD4Value = Request.Cookies["tableDataCookieHD4"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookieHD4"].Value) : "";
        //    //string cookieHD5Value = Request.Cookies["tableDataCookieHD5"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookieHD5"].Value) : "";
        //    // Kiểm tra xem có dữ liệu từ cookie không
        //    List<List<string>> combinedData = new List<List<string>>();
        //    List<List<string>> combinedDataHD = new List<List<string>>();
        //    if (!string.IsNullOrEmpty(cookie1Value) && !string.IsNullOrEmpty(cookieHD1Value) && !string.IsNullOrEmpty(cookie2Value) && !string.IsNullOrEmpty(cookieHD2Value))
        //    {
        //        // Parse chuỗi JSON thành mảng JavaScript
        //        List<List<string>> tableData = JsonConvert.DeserializeObject<List<List<string>>>(cookie1Value);
        //        List<List<string>> tableData1 = JsonConvert.DeserializeObject<List<List<string>>>(cookie2Value);
        //        List<List<string>> tableData2 = JsonConvert.DeserializeObject<List<List<string>>>(cookie3Value);
        //        //List<List<string>> tableData3 = JsonConvert.DeserializeObject<List<List<string>>>(cookie4Value);
        //        //List<List<string>> tableData4 = JsonConvert.DeserializeObject<List<List<string>>>(cookie5Value);

        //        List<List<string>> tableDataHD = JsonConvert.DeserializeObject<List<List<string>>>(cookieHD1Value);
        //        List<List<string>> tableDataHD1 = JsonConvert.DeserializeObject<List<List<string>>>(cookieHD2Value);
        //        //List<List<string>> tableDataHD2 = JsonConvert.DeserializeObject<List<List<string>>>(cookieHD3Value);
        //        //List<List<string>> tableDataHD3 = JsonConvert.DeserializeObject<List<List<string>>>(cookieHD4Value);
        //        //List<List<string>> tableDataHD4 = JsonConvert.DeserializeObject<List<List<string>>>(cookieHD5Value);
        //        //List<List<string>> tableData2 = JsonConvert.DeserializeObject<List<List<string>>>(jsonData2);
        //        combinedData.AddRange(tableData);
        //        combinedData.AddRange(tableData1);
        //        combinedData.AddRange(tableData2);
        //        //combinedData.AddRange(tableData3);
        //        //combinedData.AddRange(tableData4);

        //        combinedDataHD.AddRange(tableDataHD);
        //        combinedDataHD.AddRange(tableDataHD1);
        //        //combinedDataHD.AddRange(tableDataHD2);
        //        //combinedDataHD.AddRange(tableDataHD3);
        //        //combinedDataHD.AddRange(tableDataHD4);
        //        // Khởi tạo tệp Excel
        //        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //        using (var package = new ExcelPackage())
        //        {
        //            var worksheet = package.Workbook.Worksheets.Add("MySheet");
        //            worksheet.View.ShowGridLines = false;
        //            var startRow = 13;
        //            var startColumn = 1;
        //            // ... (Các bước tạo nội dung tệp Excel như bạn đã làm)
        //            // Đường dẫn đến hình ảnh trong thư mục 'image'
        //            var imagePath = Server.MapPath("~/assets/images/logo.png"); // Thay thế bằng đường dẫn thật
        //                                                                        // Lấy giá trị từ biến Dvcs
        //            string Dvcs = Request.Cookies["Dvcs"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs"].Value) : "";
        //            string TuNgay = Request.Cookies["tungay"] != null ? HttpUtility.UrlDecode(Request.Cookies["tungay"].Value) : "";
        //            string TuThang = Request.Cookies["tuthang"] != null ? HttpUtility.UrlDecode(Request.Cookies["tuthang"].Value) : "";
        //            string DenNgay = Request.Cookies["denngay"] != null ? HttpUtility.UrlDecode(Request.Cookies["denngay"].Value) : "";
        //            string DenThang = Request.Cookies["denthang"] != null ? HttpUtility.UrlDecode(Request.Cookies["denthang"].Value) : "";
        //            string Nam = Request.Cookies["nam"] != null ? HttpUtility.UrlDecode(Request.Cookies["nam"].Value) : "";
        //            string DiaChi = Request.Cookies["Dia_Chi"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dia_Chi"].Value) : "";
        //            string NoDauKy = Request.Cookies["NoDauKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["NoDauKy"].Value) : "";
        //            string TienHD = Request.Cookies["TienHD"] != null ? HttpUtility.UrlDecode(Request.Cookies["TienHD"].Value) : "";
        //            string TonNo = Request.Cookies["TonNo"] != null ? HttpUtility.UrlDecode(Request.Cookies["TonNo"].Value) : "";
        //            string TienChu = Request.Cookies["TienChu"] != null ? HttpUtility.UrlDecode(Request.Cookies["TienChu"].Value) : "";
        //            string TonNo2 = Request.Cookies["TonNo2"] != null ? HttpUtility.UrlDecode(Request.Cookies["TonNo2"].Value) : "";
        //            string NgayTT = Request.Cookies["NgayTT"] != null ? HttpUtility.UrlDecode(Request.Cookies["NgayTT"].Value) : "";
        //            string ChiNhanh = Request.Cookies["ChiNhanh"] != null ? HttpUtility.UrlDecode(Request.Cookies["ChiNhanh"].Value) : "";
        //            string DiaChi2 = Request.Cookies["DiaChi2"] != null ? HttpUtility.UrlDecode(Request.Cookies["DiaChi2"].Value) : "";
        //            //string Dvcs1 = Request.Cookies["Dvcs1"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs1"].Value) : "";
        //            string ten_kh = Request.Cookies["Ten_Dt"] != null ? HttpUtility.UrlDecode(Request.Cookies["Ten_Dt"].Value) : "";
        //            string NgayKy = Request.Cookies["NgayKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["NgayKy"].Value) : "";
        //            string ThangKy = Request.Cookies["ThangKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["ThangKy"].Value) : "";
        //            string NamKy = Request.Cookies["NamKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["NamKy"].Value) : "";
        //            //string HanNgay = Request.Cookies["HanNgay"] != null ? HttpUtility.UrlDecode(Request.Cookies["HanNgay"].Value) : "";
        //            string CN = Request.Cookies["CN"] != null ? HttpUtility.UrlDecode(Request.Cookies["CN"].Value) : "";
        //            string TK = Request.Cookies["TK"] != null ? HttpUtility.UrlDecode(Request.Cookies["TK"].Value) : "";
        //            string LH = Request.Cookies["LH"] != null ? HttpUtility.UrlDecode(Request.Cookies["LH"].Value) : "";

        //            string TienTT = Request.Cookies["Tien_TT"] != null ? HttpUtility.UrlDecode(Request.Cookies["Tien_TT"].Value) : "";
        //            string CKTT = Request.Cookies["CKTT"] != null ? HttpUtility.UrlDecode(Request.Cookies["CKTT"].Value) : "";
        //            string TCCK = Request.Cookies["TC_CKTT_TienTT"] != null ? HttpUtility.UrlDecode(Request.Cookies["TC_CKTT_TienTT"].Value) : "";
        //            // Đặt font chữ "Arial" cho toàn bộ tệp Excel
        //            worksheet.Cells.Style.Font.Name = "Times New Roman";

        //            // Chèn hình ảnh từ tệp hình vào ô A1
        //            ExcelPicture picture = worksheet.Drawings.AddPicture("MyPicture", new FileInfo(imagePath));
        //            picture.SetSize(70, 50); // Đặt kích thước cho hình ảnh
        //            picture.From.Row = 1;
        //            picture.From.Column = 0;

        //            worksheet.Column(1).Width = 8;

        //            // Đặt văn bản vào ô A2
        //            worksheet.Cells["B1"].Value = "CTY CỔ PHẦN DƯỢC PHẨM OPC";
        //            var cellB1 = worksheet.Cells["B1"];
        //            cellB1.Style.Font.Bold = true;
        //            //worksheet.Cells["B1"].Style.Indent = 3;
        //            //worksheet.Cells["B2"].Style.Indent = 3;
        //            //worksheet.Cells["B3"].Style.Indent = 3;
        //            worksheet.Cells["B2"].Value = Dvcs;
        //            worksheet.Cells["B3"].Value = $"Số:";
        //            worksheet.Cells["H1"].Value = "Cộng Hòa Xã Hội Chủ Nghĩa Việt Nam";
        //            worksheet.Cells["H2"].Value = "Độc Lập - Tự Do - Hạnh Phúc";
        //            worksheet.Cells["H2"].Style.Indent = 4;

        //            worksheet.Cells["E4"].Value = "BẢNG ĐỐI CHIẾU DOANH THU CÔNG NỢ";
        //            worksheet.Cells["E4"].Style.Font.Bold = true;
        //            worksheet.Cells["E4"].Style.Font.Size = 16;
        //            worksheet.Cells["E5"].Value = $"Từ ngày {TuNgay} tháng {TuThang} đến ngày {DenNgay} tháng {DenThang} năm {Nam}";
        //            worksheet.Cells["E5"].Style.Indent = 4;
        //            worksheet.Cells["A7"].Value = $"Tên khách hàng: {ten_kh}";
        //            worksheet.Cells["A7"].Style.Font.Bold = true;
        //            worksheet.Cells["A8"].Value = $"Địa chỉ khách hàng: {DiaChi}";
        //            worksheet.Cells["A8"].Style.Font.Bold = true;
        //            worksheet.Cells["A9"].Value = $"I.Số dư nợ trước ngày: {TuNgay}/{TuThang}/{Nam}";
        //            worksheet.Cells["A9"].Style.Font.Bold = true;
        //            worksheet.Cells["E9"].Value = $"mang sang {NoDauKy} đồng";
        //            worksheet.Cells["E9"].Style.Font.Bold = true;
        //            worksheet.Cells["A10"].Value = "II.Doanh thu và công nợ phát sinh trong kỳ đối chiếu này: ";
        //            worksheet.Cells["A10"].Style.Font.Bold = true;
        //            //worksheet.Cells["A8"].Value = $"{Dvcs} - Công ty Cổ Phần Dược Phẩm OPC trân trọng thông báo đến quý khách hàng có số dư nợ mà Quý Khách";
        //            //worksheet.Cells["A9"].Value = $"hàng chưa thanh toán cho chúng tôi tính đến ngày {denngay} là: {tongno}";
        //            //worksheet.Cells["B11"].Value = $"Trong đó nợ quá hạn là: {QuaHan} bao gồm các hóa đơn sau:";

        //            var sttCell = worksheet.Cells[startRow - 1, startColumn, startRow, startColumn];
        //            sttCell.Merge = true;
        //            sttCell.Value = "STT";
        //            sttCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            sttCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
        //            sttCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc
        //            sttCell.Style.Font.Bold = true;
        //            sttCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            var KHMua = worksheet.Cells[startRow - 1, startColumn + 1, startRow - 1, startColumn + 3];
        //            KHMua.Merge = true;
        //            KHMua.Value = "KHÁCH HÀNG MUA";
        //            KHMua.Style.Font.Bold = true;
        //            KHMua.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            KHMua.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            KHMua.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow, startColumn + 1].Value = "SỐ";
        //            worksheet.Cells[startRow, startColumn + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow, startColumn + 1].Style.Font.Bold = true;
        //            worksheet.Cells[startRow, startColumn + 2].Value = "NGÀY";
        //            worksheet.Cells[startRow, startColumn + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow, startColumn + 2].Style.Font.Bold = true;
        //            worksheet.Cells[startRow, startColumn + 3].Value = "SỐ TIỀN";
        //            worksheet.Cells[startRow, startColumn + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow, startColumn + 3].Style.Font.Bold = true;
        //            var KHTT = worksheet.Cells[startRow - 1, startColumn + 4, startRow - 1, startColumn + 8];
        //            KHTT.Merge = true;
        //            KHTT.Value = "KHÁCH HÀNG THANH TOÁN/TRẢ HÀNG BÙ TRỪ";
        //            KHTT.Style.Font.Bold = true;
        //            KHTT.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            //worksheet.Column(startColumn + 4).Width = 20;
        //            worksheet.Cells[startRow - 1, startColumn + 9].Value = "";
        //            worksheet.Cells[startRow - 1, startColumn + 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow, startColumn + 4].Value = "SỐ";
        //            worksheet.Cells[startRow, startColumn + 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow, startColumn + 4].Style.Font.Bold = true;

        //            worksheet.Cells[startRow, startColumn + 5].Value = "NGÀY";
        //            worksheet.Cells[startRow, startColumn + 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow, startColumn + 5].Style.Font.Bold = true;
        //            worksheet.Cells[startRow, startColumn + 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        //            worksheet.Cells[startRow, startColumn + 6].Value = "SỐ TIỀN";
        //            worksheet.Cells[startRow, startColumn + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow, startColumn + 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 6].Style.Font.Bold = true;

        //            worksheet.Cells[startRow, startColumn + 7].Value = "CKTT";
        //            worksheet.Cells[startRow, startColumn + 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow, startColumn + 7].Style.Font.Bold = true;

        //            worksheet.Cells[startRow, startColumn + 8].Value = "TỔNG TIỀN";
        //            worksheet.Cells[startRow, startColumn + 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow, startColumn + 8].Style.Font.Bold = true;
        //            worksheet.Cells[startRow, startColumn + 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        //            worksheet.Cells[startRow, startColumn + 9].Value = "GHI CHÚ";
        //            worksheet.Cells[startRow, startColumn + 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            worksheet.Cells[startRow, startColumn + 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow, startColumn + 9].Style.Font.Bold = true;
        //            if (combinedData != null && combinedData.Any())
        //            {
        //                // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
        //                for (int row = 0; row < combinedData.Count - 1; row++)
        //                {
        //                    var rowData = combinedData[row];
        //                    for (int col = 0; col < rowData.Count; col++)
        //                    {
        //                        if (col == 4 && col == 7 && col == 9)
        //                        {
        //                            worksheet.Cells[startRow - 1 + row, startColumn + col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        //                            worksheet.Cells[startRow - 1 + row, startColumn + col].Value = rowData[col];
        //                            worksheet.Cells[startRow - 1 + row, startColumn + col].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //                        }
        //                        else
        //                        {
        //                            worksheet.Cells[startRow - 1 + row, startColumn + col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //                            worksheet.Cells[startRow - 1 + row, startColumn + col].Value = rowData[col];
        //                            worksheet.Cells[startRow - 1 + row, startColumn + col].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //                        }


        //                    }
        //                }
        //            }
        //            else
        //            {
        //                worksheet.Cells[startRow, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
        //            }
        //            var TC = worksheet.Cells[startRow - 2 + combinedData.Count, startColumn, startRow - 2 + combinedData.Count, startColumn + 2];
        //            TC.Merge = true;
        //            TC.Value = "Tổng cộng: ";
        //            TC.Style.Font.Bold = true;
        //            TC.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            TC.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
        //            TC.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 3].Value = $"{TienHD}"; // Ví dụ: Ghi giá trị tổng vào cột thứ 4
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 3].Style.Font.Bold = true;
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 6].Value = $"{TienTT}";
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 7].Value = $"{CKTT}";
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 8].Value = $"{TCCK}";
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 6].Style.Font.Bold = true;
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 7].Style.Font.Bold = true;
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 8].Style.Font.Bold = true;
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRow - 2 + combinedData.Count, startColumn + 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

        //            var endII = startRow + combinedData.Count;
        //            var nextII = endII + 1;
        //            worksheet.Cells[nextII, startColumn].Value = $"III. Số tiền khách hàng chưa thanh toán, tính đến cuối ngày: {DenNgay}/{DenThang}/{Nam} là: {TonNo} đồng";
        //            worksheet.Cells[nextII, startColumn].Style.Font.Bold = true;
        //            worksheet.Cells[nextII + 1, startColumn].Value = $"Số tiền bằng chữ: {TienChu}";
        //            worksheet.Cells[nextII + 1, startColumn].Style.Font.Bold = true;
        //            worksheet.Cells[nextII + 2, startColumn].Value = "Chi tiết các hóa đơn chưa thanh toán: ";

        //            var startRowIII = nextII + 3;

        //            var sttIIICell = worksheet.Cells[startRowIII + 1, startColumn, startRowIII + 2, startColumn + 1];
        //            sttIIICell.Merge = true;
        //            worksheet.Column(startColumn).Width = 15; // Đặt chiều rộng của cột chứa ô "STT" thành 15 đơn vị.

        //            sttIIICell.Value = "STT";
        //            sttIIICell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
        //            sttIIICell.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc
        //            sttIIICell.Style.Font.Bold = true;
        //            sttIIICell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

        //            var HD = worksheet.Cells[startRowIII + 1, startColumn + 2, startRowIII + 1, startColumn + 9];
        //            HD.Merge = true;
        //            HD.Value = "HÓA ĐƠN";
        //            HD.Style.Font.Bold = true;
        //            HD.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            HD.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            HD.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

        //            var SO = worksheet.Cells[startRowIII + 2, startColumn + 2, startRowIII + 2, startColumn + 3];
        //            SO.Merge = true;
        //            SO.Value = "SỐ";
        //            worksheet.Column(startColumn + 1).Width = 15;
        //            SO.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            SO.Style.Font.Bold = true;
        //            SO.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            SO.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

        //            var NGAY = worksheet.Cells[startRowIII + 2, startColumn + 4, startRowIII + 2, startColumn + 5];
        //            NGAY.Merge = true;
        //            NGAY.Value = "NGÀY";
        //            worksheet.Column(startColumn + 2).Width = 15;
        //            NGAY.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            NGAY.Style.Font.Bold = true;
        //            NGAY.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            NGAY.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

        //            var TIENHD = worksheet.Cells[startRowIII + 2, startColumn + 6, startRowIII + 2, startColumn + 7];
        //            TIENHD.Merge = true;
        //            TIENHD.Value = "SỐ TIỀN HD";
        //            worksheet.Column(startColumn + 3).Width = 15;
        //            TIENHD.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            TIENHD.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            TIENHD.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            TIENHD.Style.Font.Bold = true;

        //            var GHICHU = worksheet.Cells[startRowIII + 2, startColumn + 8, startRowIII + 2, startColumn + 9];
        //            GHICHU.Merge = true;
        //            GHICHU.Value = "GHI CHÚ";
        //            worksheet.Column(startColumn + 4).Width = 15;
        //            worksheet.Column(startColumn + 5).Width = 15;
        //            worksheet.Column(startColumn + 6).Width = 15;
        //            worksheet.Column(startColumn + 8).Width = 15;
        //            GHICHU.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            GHICHU.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            GHICHU.Style.Font.Bold = true;
        //            GHICHU.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        //            if (combinedDataHD != null && combinedDataHD.Any())
        //            {
        //                // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
        //                for (int row = 0; row < combinedDataHD.Count - 1; row++)
        //                {
        //                    var rowData = combinedDataHD[row];
        //                    for (int col = 0; col < rowData.Count; col++)
        //                    {
        //                        //var GOP = worksheet.Cells[startRowIII + 1 + row, startColumn + col, startRowIII + 1 + row, startColumn + col+ 1];
        //                        //GOP.Merge = true;
        //                        //GOP.Value = rowData[col];
        //                        //GOP.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //                        var cell = worksheet.Cells[startRowIII + 1 + row, startColumn + col * 2];
        //                        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //                        cell.Value = rowData[col];

        //                        var mergeCell = worksheet.Cells[startRowIII + 1 + row, startColumn + col * 2, startRowIII + 1 + row, startColumn + col * 2 + 1];
        //                        mergeCell.Merge = true;
        //                        mergeCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                worksheet.Cells[startRowIII, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
        //            }
        //            var TC2 = worksheet.Cells[startRowIII + combinedDataHD.Count, startColumn, startRowIII + combinedDataHD.Count, startColumn + 5];
        //            TC2.Merge = true;
        //            TC2.Value = "Tổng cộng: ";
        //            TC2.Style.Font.Bold = true;
        //            TC2.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            TC2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
        //            TC2.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc

        //            var TONNO2 = worksheet.Cells[startRowIII + combinedDataHD.Count, startColumn + 6, startRowIII + combinedDataHD.Count, startColumn + 7];
        //            TONNO2.Merge = true;
        //            TONNO2.Value = $"{TonNo2}";
        //            //worksheet.Cells[startRowIII + 1 + combinedDataHD.Count, startColumn + 6].Value = $"{TonNo2}";
        //            TONNO2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        //            TONNO2.Style.Font.Bold = true;
        //            TONNO2.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            worksheet.Cells[startRowIII + combinedDataHD.Count, startColumn + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

        //            var lastCell = worksheet.Cells[startRowIII + combinedDataHD.Count, startColumn + 8, startRowIII + combinedDataHD.Count, startColumn + 9];
        //            lastCell.Merge = true;
        //            lastCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        //            var endRowIII = startRowIII + 1 + combinedDataHD.Count + 1;
        //            worksheet.Cells[endRowIII, startColumn].Value = $"Xin vui lòng xác nhận và gửi lại cho {Dvcs} trước ngày {NgayTT}";
        //            worksheet.Cells[endRowIII, startColumn].Style.Font.Bold = true;
        //            worksheet.Cells[endRowIII + 1, startColumn].Value = $"Nơi nhận: {ChiNhanh}";
        //            worksheet.Cells[endRowIII + 1, startColumn].Style.Font.Bold = true;
        //            worksheet.Cells[endRowIII + 2, startColumn].Value = $"Địa chỉ: {DiaChi2}";
        //            worksheet.Cells[endRowIII + 2, startColumn].Style.Font.Bold = true;
        //            worksheet.Cells[endRowIII + 3, startColumn].Value = $"Khi cần đối chiếu số liệu liên hệ: {LH}";
        //            worksheet.Cells[endRowIII + 3, startColumn].Style.Font.Bold = true;
        //            worksheet.Cells[endRowIII + 4, startColumn].Value = $"Số tiền còn nợ đề nghị Quý khách hàng thanh toán bằng tiền mặt hoặc chuyển khoản vào tài khoản {CN}, số";
        //            worksheet.Cells[endRowIII + 4, startColumn].Style.Font.Bold = true;
        //            worksheet.Cells[endRowIII + 5, startColumn].Value = $"tài khoản: {TK}";
        //            worksheet.Cells[endRowIII + 5, startColumn].Style.Font.Bold = true;
        //            worksheet.Cells[endRowIII + 7, startColumn].Value = "Trân trọng cảm ơn!";
        //            worksheet.Cells[endRowIII + 7, startColumn].Style.Font.Bold = true;
        //            worksheet.Cells[endRowIII + 7, startColumn].Style.Font.Italic = true;
        //            worksheet.Cells[endRowIII + 8, startColumn + 7].Value = $"Ngày {NgayKy} tháng {ThangKy} năm {NamKy}";
        //            worksheet.Cells[endRowIII + 8, startColumn + 7].Style.Font.Bold = true;
        //            worksheet.Cells[endRowIII + 9, startColumn].Value = "ĐẠI DIỆN KHÁCH HÀNG";
        //            worksheet.Cells[endRowIII + 9, startColumn].Style.Font.Bold = true;
        //            worksheet.Cells[endRowIII + 9, startColumn + 7].Value = "ĐẠI DIỆN CHI NHÁNH";
        //            worksheet.Cells[endRowIII + 9, startColumn + 7].Style.Font.Bold = true;
        //            worksheet.Cells[endRowIII + 9, startColumn + 7].Style.Indent = 2;



        //            package.Save();
        //            byte[] fileBytes = package.GetAsByteArray();

        //            // Trả về tệp Excel dưới dạng dữ liệu binary
        //            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

        //        }


        //    }
        //    else
        //    {
        //        return Content("Không có dữ liệu từ cookie.");
        //    }


        //}
        public ActionResult ExportPhieuXuatKho()
        {
            var fileName = $"PhieuXuatKho{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            // Lấy dữ liệu từ cookie
            string jsonData = Request.Cookies["tableDataCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookie"].Value) : "";

            // Kiểm tra xem có dữ liệu từ cookie không
            if (!string.IsNullOrEmpty(jsonData))
            {
                // Parse chuỗi JSON thành mảng JavaScript
                List<List<string>> tableData = JsonConvert.DeserializeObject<List<List<string>>>(jsonData);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("MySheet");
                    worksheet.View.ShowGridLines = false;

                    // ... (Các bước tạo nội dung tệp Excel như bạn đã làm)
                    // Đường dẫn đến hình ảnh trong thư mục 'image'
                    /*  var imagePath = Server.MapPath("~/assets/images/logo.png");*/ // Thay thế bằng đường dẫn thật
                                                                                      // Lấy giá trị từ biến Dvcs
                    string Ngay = Request.Cookies["ngay"] != null ? HttpUtility.UrlDecode(Request.Cookies["ngay"].Value) : "";
                    string SoCT = Request.Cookies["SoCt"] != null ? HttpUtility.UrlDecode(Request.Cookies["SoCt"].Value) : "";
                    string Thang = Request.Cookies["thang"] != null ? HttpUtility.UrlDecode(Request.Cookies["thang"].Value) : "";
                    string Nam = Request.Cookies["nam"] != null ? HttpUtility.UrlDecode(Request.Cookies["nam"].Value) : "";
                    string DVNH = Request.Cookies["Ten_dt"] != null ? HttpUtility.UrlDecode(Request.Cookies["Ten_dt"].Value) : "";
                    string QuaHan = Request.Cookies["QuaHan"] != null ? HttpUtility.UrlDecode(Request.Cookies["QuaHan"].Value) : "";
                    string HanNgay = Request.Cookies["HanNgay"] != null ? HttpUtility.UrlDecode(Request.Cookies["HanNgay"].Value) : "";
                    string CN = Request.Cookies["CN"] != null ? HttpUtility.UrlDecode(Request.Cookies["CN"].Value) : "";
                    string TK = Request.Cookies["TK"] != null ? HttpUtility.UrlDecode(Request.Cookies["TK"].Value) : "";
                    string LH = Request.Cookies["LH"] != null ? HttpUtility.UrlDecode(Request.Cookies["LH"].Value) : "";
                    // Đặt font chữ "Arial" cho toàn bộ tệp Excel
                    worksheet.Cells.Style.Font.Name = "Times New Roman";

                    // Đặt văn bản vào ô A2
                    worksheet.Cells["A1"].Value = "CTY CỔ PHẦN DƯỢC PHẨM OPC";
                    worksheet.Cells["A1"].Style.Font.Bold = true;
                    var cellB1 = worksheet.Cells["B1"];
                    cellB1.Style.Font.Bold = true;
                    worksheet.Cells["C3"].Value = "PHIẾU XUẤT KHO CỬA HÀNG QUẬN 10";
                    worksheet.Cells["C3"].Style.Font.Bold = true;
                    worksheet.Cells["C3"].Style.Font.Size = 16;

                    worksheet.Cells["G4"].Value = $"Số: {SoCT}";
                    worksheet.Cells["G4"].Style.Font.Bold = true;

                    //worksheet.Cells["F4"].Value = $"hàng chưa thanh toán cho chúng tôi tính đến ngày {denngay} là: {tongno}";
                    worksheet.Cells["D5"].Value = $"Ngày {Ngay} tháng {Thang} năm {Nam}";
                    worksheet.Cells["A6"].Value = $"Đơn vị xuất hàng: Kho thành phẩm Cửa hàng Quận 10";
                    worksheet.Cells["A7"].Value = "Địa chỉ: 134/1 Tô Hiến Thành, P15, Quận 10, TP.HCM";
                    worksheet.Cells["A8"].Value = $"Đơn vị nhận hàng: {DVNH}";
                    worksheet.Cells["A9"].Value = $"Diễn giải: Xuất hàng giao cho khách";
                    var startRow = 13;
                    var startColumn = 1;
                    worksheet.Cells[startRow - 1, startColumn].Value = "STT";
                    worksheet.Cells[startRow - 1, startColumn + 1].Value = "TÊN SẢN PHẨM - QUY CÁCH";
                    worksheet.Cells[startRow - 1, startColumn + 2].Value = "DVT";
                    worksheet.Cells[startRow - 1, startColumn + 3].Value = "SỐ LƯỢNG";

                    for (int col = 0; col < 4; col++)
                    {
                        var columnHeaderCell = worksheet.Cells[startRow - 1, startColumn + col];
                        columnHeaderCell.Style.Font.Bold = true;
                        columnHeaderCell.Style.Font.Size = 10;
                        columnHeaderCell.Style.Font.Color.SetColor(Color.Black);
                        columnHeaderCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        columnHeaderCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        columnHeaderCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        columnHeaderCell.Style.Fill.BackgroundColor.SetColor(Color.White);
                    }
                    var columnHeaderStyle = worksheet.Cells[startRow - 1, startColumn, startRow - 1, startColumn + 3].Style;
                    columnHeaderStyle.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black); // Đóng khung solid đen
                    worksheet.Column(startColumn).Width = 10; // Độ rộng cột cho "STT"
                    worksheet.Column(startColumn + 1).Width = 25; // Độ rộng cột cho "SỐ HÓA ĐƠN"
                    worksheet.Column(startColumn + 2).Width = 15; // Độ rộng cột cho "NGÀY XUẤT"
                    worksheet.Column(startColumn + 3).Width = 15; // Độ rộng cột cho "TIỀN NỢ"


                    // Đảm bảo rằng có dữ liệu trong biến tableData
                    if (tableData != null && tableData.Any())
                    {
                        // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
                        for (int row = 0; row < tableData.Count; row++)
                        {
                            var rowData = tableData[row];
                            for (int col = 0; col < rowData.Count; col++)
                            {
                                worksheet.Cells[startRow + row, startColumn + col].Value = rowData[col];
                                worksheet.Cells[startRow + row, startColumn + col].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                                worksheet.Cells[startRow + row, startColumn + col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[startRow + row, startColumn + col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            }
                        }
                    }
                    else
                    {
                        worksheet.Cells[startRow, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
                    }
                    worksheet.Cells[startRow + tableData.Count, startColumn + 1].Value = "Tổng cộng";
                    worksheet.Cells[startRow + tableData.Count, startColumn + 1].Style.Font.Bold = true;
                    worksheet.Cells[startRow + tableData.Count, startColumn].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow + tableData.Count, startColumn + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow + tableData.Count, startColumn + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow + tableData.Count, startColumn + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    int defaultHeaderRowIndex = 13;
                    // Xóa hàng tiêu đề mặc định
                    worksheet.DeleteRow(defaultHeaderRowIndex);
                    //var dataRowStyle = worksheet.Cells[startRow, startColumn, startRow, startColumn + 5].Style;
                    //dataRowStyle.Font.Bold = false;
                    //dataRowStyle.Font.Color.SetColor(Color.Black);
                    //dataRowStyle.Fill.PatternType = ExcelFillStyle.None;
                    // Tạo bảng trong tệp Excel
                    var endRow = startRow + tableData.Count;
                    var endColumn = 6;
                    worksheet.DeleteRow(endRow, 1);



                    int nextRow = endRow + 1;
                    //worksheet.Cells[nextRow, startColumn].Value = $"Kính đề nghị Quý khách vui lòng đối chiếu và xác nhận số tiền gửi về {Dvcs} - Công Ty Cổ Phần Dược Phẩm OPC";
                    worksheet.Cells[nextRow + 1, startColumn + 5].Value = $"Ngày     tháng     năm";
                    worksheet.Cells[nextRow + 1, startColumn + 5].Style.Font.Bold = true;
                    worksheet.Cells[nextRow + 2, startColumn + 1].Value = $"Bên nhận";
                    worksheet.Cells[nextRow + 2, startColumn + 1].Style.Font.Bold = true;
                    worksheet.Cells[nextRow + 2, startColumn + 4].Value = $"Người lập phiếu";
                    worksheet.Cells[nextRow + 2, startColumn + 4].Style.Indent = 7;
                    worksheet.Cells[nextRow + 2, startColumn + 4].Style.Font.Bold = true;
                    worksheet.Cells[nextRow + 3, startColumn + 1].Value = "(Ký, họ tên)";
                    worksheet.Cells[nextRow + 3, startColumn + 1].Style.Font.Bold = true;
                    worksheet.Cells[nextRow + 3, startColumn + 1].Style.Font.Italic = true;

                    worksheet.Cells[nextRow + 3, startColumn + 4].Value = "(Ký, họ tên)";
                    worksheet.Cells[nextRow + 3, startColumn + 4].Style.Font.Bold = true;
                    worksheet.Cells[nextRow + 3, startColumn + 4].Style.Font.Italic = true;
                    worksheet.Cells[nextRow + 3, startColumn + 4].Style.Indent = 7;
                    //worksheet.Cells[nextRow + 4, startColumn].Value = $"Khi cần đối chiếu xin liên hệ {LH}";
                    //worksheet.Cells[nextRow + 4, startColumn].Style.Indent = 2;
                    //worksheet.Cells[nextRow + 6, startColumn].Value = "Trân trọng!";
                    //worksheet.Cells[nextRow + 6, startColumn].Style.Indent = 2;
                    //worksheet.Cells[nextRow + 6, startColumn].Style.Font.Italic = true;
                    //worksheet.Cells[nextRow + 8, startColumn + 1].Value = "Khách Hàng Xác Nhận";
                    //worksheet.Cells[nextRow + 8, startColumn + 1].Style.Font.Bold = true;
                    //worksheet.Cells[nextRow + 8, startColumn + 4].Value = "Giám Đốc";
                    //worksheet.Cells[nextRow + 8, startColumn + 4].Style.Font.Bold = true;
                    //worksheet.Cells[nextRow + 8, startColumn + 7].Value = "Kế Toán";
                    //worksheet.Cells[nextRow + 8, startColumn + 7].Style.Font.Bold = true;
                    //worksheet.Cells[nextRow + 9, startColumn].Value = "(Ký, đóng dấu, ghi rõ họ tên)";
                    //worksheet.Cells[nextRow + 9, startColumn].Style.Indent = 4;
                    //worksheet.Cells[nextRow + 9, startColumn].Style.Font.Italic = true;

                    package.Save();
                    byte[] fileBytes = package.GetAsByteArray();

                    // Trả về tệp Excel dưới dạng dữ liệu binary
                    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

                }


            }
            else
            {
                return Content("Không có dữ liệu từ cookie.");
            }

        }
        public List<GetData> LoadDataDTCN()
        {
            connectSQL();
            List<GetData> dataItems = new List<GetData>();
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DoiChieuDoanhThuCongNo_SAP]", connection))
                {
                    var fromDate = Request.Cookies["From_date"].Value;
                    var toDate = Request.Cookies["To_Date"].Value;
                    var NgayTT = Request.Cookies["Ngay_TT"].Value;
                    var ma_dvcs = Request.Cookies["MA_DVCS"].Value;
                    var ma_dt = Request.Cookies["Ma_Dt"].Value;
                    var NgayKy = Request.Cookies["Ngay_Ky"].Value;
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;
                    using (SqlDataAdapter sda = new SqlDataAdapter(command))
                    {
                        DataSet ds = new DataSet();
                        command.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                        command.Parameters.AddWithValue("@_Den_Ngay", toDate);
                        command.Parameters.AddWithValue("@_Ma_DvCs", ma_dvcs);
                        command.Parameters.AddWithValue("@_Ngay_TT", NgayTT);
                        command.Parameters.AddWithValue("@_Ma_Dt", ma_dt);
                        command.Parameters.AddWithValue("@_Ngay_Ky", NgayKy);
                        sda.Fill(ds);

                        // Kiểm tra xem DataSet có bảng dữ liệu hay không
                        if (ds.Tables.Count > 0)
                        {
                            DataTable dt = ds.Tables[2];

                            foreach (DataRow row in dt.Rows)
                            {
                                GetData dataItem = new GetData
                                {
                                    So = row["So_Ct"].ToString(),
                                    Ngay = row["Ngay_Ct1"].ToString(),
                                    GhiChu = row["Ghi_Chu"].ToString(),
                                    TienHD = row["Ton_No1"].ToString(),
                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }
            return dataItems;
        }
        public List<GetData> LoadDataDTCN2()
        {
            connectSQL();
            List<GetData> dataItems = new List<GetData>();
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DoiChieuDoanhThuCongNo_SAP]", connection))
                {
                    var fromDate = Request.Cookies["From_date"].Value;
                    var toDate = Request.Cookies["To_Date"].Value;
                    var NgayTT = Request.Cookies["Ngay_TT"].Value;
                    var ma_dvcs = Request.Cookies["MA_DVCS"].Value;
                    var ma_dt = Request.Cookies["MaDT"].Value;
                    var NgayKy = Request.Cookies["Ngay_Ky"].Value;
                    if (ma_dvcs == "OPC_B1")
                    {
                        string ma_dvcsFirst3Chars = ma_dvcs == "OPC_B1" ? ma_dvcs.Substring(0, 3) : ma_dvcs;
                        string first3Chars = ma_dvcsFirst3Chars.Substring(0, 3);
                        ma_dvcs = first3Chars;
                    }
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;
                    using (SqlDataAdapter sda = new SqlDataAdapter(command))
                    {
                        DataSet ds = new DataSet();
                        command.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                        command.Parameters.AddWithValue("@_Den_Ngay", toDate);
                        command.Parameters.AddWithValue("@_Ma_DvCs", ma_dvcs);
                        command.Parameters.AddWithValue("@_Ngay_TT", NgayTT);
                        command.Parameters.AddWithValue("@_Ma_Dt", ma_dt);
                        command.Parameters.AddWithValue("@_Ngay_Ky", NgayKy);
                        sda.Fill(ds);

                        // Kiểm tra xem DataSet có bảng dữ liệu hay không
                        if (ds.Tables.Count > 0)
                        {
                            DataTable dt = ds.Tables[2];

                            foreach (DataRow row in dt.Rows)
                            {
                                GetData dataItem = new GetData
                                {
                                    So = row["So_Ct"].ToString(),
                                    Ngay = row["Ngay_Ct1"].ToString(),
                                    GhiChu = row["Ghi_Chu"].ToString(),
                                    TienHD = row["Ton_No1"].ToString(),
                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }
            return dataItems;
        }
        public ActionResult BangDoiChieuDTCN_Fill()
        {
            List<MauInChungTu> dmDlist = LoadDmDt("");

            ViewBag.DataItems = dmDlist;
            return View();
        }
        public ActionResult BangDoiChieuDTCN(MauInChungTu MauIn)
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;

            if (ma_dvcs == "OPC_B1")
            {
                string ma_dvcsFirst3Chars = ma_dvcs == "OPC_B1" ? ma_dvcs.Substring(0, 3) : ma_dvcs;
                string first3Chars = ma_dvcsFirst3Chars.Substring(0, 3);
                ma_dvcs = first3Chars;
            }


            DataSet ds = new DataSet();
            connectSQL();

            //MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_DoiChieuDoanhThuCongNo_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                var MaDT = Request.Cookies["Ma_DT"].Value;
                //MauIn.From_date = Request.Cookies["From_date"].Value;
                //MauIn.To_date = Request.Cookies["To_Date"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {


                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_Ma_dt", MaDT);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                    cmd.Parameters.AddWithValue("@_Ngay_TT", MauIn.Ngay_TT);
                    cmd.Parameters.AddWithValue("@_Ngay_Ky", MauIn.Ngay_Ky);
                    cmd.Parameters.AddWithValue("@_So", MauIn.So);


                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult ExportDoiChieuCongNo()
        {
            var fileName = $"BangDoiChieuCongNo{DateTime.Now:yyyyMMddHHmmss}.xlsx";


            List<GetData> combinedData = LoadDataDTCN();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var imagePath = Server.MapPath("~/assets/images/logo.png");
                var worksheet = package.Workbook.Worksheets.Add("MySheet");
                worksheet.View.ShowGridLines = false;

                // ... (Các bước tạo nội dung tệp Excel như bạn đã làm)
                // Đường dẫn đến hình ảnh trong thư mục 'image'
                /*  var imagePath = Server.MapPath("~/assets/images/logo.png");*/ // Thay thế bằng đường dẫn thật
                string Dvcs = Request.Cookies["Dvcs"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs"].Value) : "";
                string Dvcs1 = Request.Cookies["Dvcs1"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs1"].Value) : "";// Lấy giá trị từ biến Dvcs
                string ten_dt = Request.Cookies["ten_dt"] != null ? HttpUtility.UrlDecode(Request.Cookies["ten_dt"].Value) : "";
                string TruocNgay = Request.Cookies["TruocNgay"] != null ? HttpUtility.UrlDecode(Request.Cookies["TruocNgay"].Value) : "";
                string TruocThang = Request.Cookies["TruocThang"] != null ? HttpUtility.UrlDecode(Request.Cookies["TruocThang"].Value) : "";
                string TruocNam = Request.Cookies["TruocNam"] != null ? HttpUtility.UrlDecode(Request.Cookies["TruocNam"].Value) : "";
                string DenNgay = Request.Cookies["DenNgay"] != null ? HttpUtility.UrlDecode(Request.Cookies["DenNgay"].Value) : "";
                string DenThang = Request.Cookies["DenThang"] != null ? HttpUtility.UrlDecode(Request.Cookies["DenThang"].Value) : "";
                string DenNam = Request.Cookies["DenNam"] != null ? HttpUtility.UrlDecode(Request.Cookies["DenNam"].Value) : "";
                string Nam = Request.Cookies["nam"] != null ? HttpUtility.UrlDecode(Request.Cookies["nam"].Value) : "";
                string NoDauKy = Request.Cookies["NoDauKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["NoDauKy"].Value) : "";
                string TonNo = Request.Cookies["TonNo"] != null ? HttpUtility.UrlDecode(Request.Cookies["TonNo"].Value) : "";
                string TienHD = Request.Cookies["TienHD"] != null ? HttpUtility.UrlDecode(Request.Cookies["TienHD"].Value) : "";
                string TienTT = Request.Cookies["TienTT"] != null ? HttpUtility.UrlDecode(Request.Cookies["TienTT"].Value) : "";
                string TienChu = Request.Cookies["TienChu"] != null ? HttpUtility.UrlDecode(Request.Cookies["TienChu"].Value) : "";
                string TonNo2 = Request.Cookies["TonNo2"] != null ? HttpUtility.UrlDecode(Request.Cookies["TonNo2"].Value) : "";
                string DiaChi = Request.Cookies["DiaChi2"] != null ? HttpUtility.UrlDecode(Request.Cookies["DiaChi2"].Value) : "";
                string NgayTT = Request.Cookies["NgayTT"] != null ? HttpUtility.UrlDecode(Request.Cookies["NgayTT"].Value) : "";
                string CN = Request.Cookies["CN"] != null ? HttpUtility.UrlDecode(Request.Cookies["CN"].Value) : "";
                string SDT = Request.Cookies["SDT"] != null ? HttpUtility.UrlDecode(Request.Cookies["SDT"].Value) : "";
                string TK = Request.Cookies["TK"] != null ? HttpUtility.UrlDecode(Request.Cookies["TK"].Value) : "";
                string LH = Request.Cookies["LH"] != null ? HttpUtility.UrlDecode(Request.Cookies["LH"].Value) : "";
                string NgayKy = Request.Cookies["NgayKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["NgayKy"].Value) : "";
                string ThangKy = Request.Cookies["ThangKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["ThangKy"].Value) : "";
                string NamKy = Request.Cookies["NamKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["NamKy"].Value) : "";
                string So = Request.Cookies["so"] != null ? HttpUtility.UrlDecode(Request.Cookies["so"].Value) : "";
                string Time = Request.Cookies["Time"] != null ? HttpUtility.UrlDecode(Request.Cookies["Time"].Value) : "";
                // Đặt font chữ "Arial" cho toàn bộ tệp Excel
                worksheet.Cells.Style.Font.Name = "Times New Roman";
                ExcelPicture picture = worksheet.Drawings.AddPicture("MyPicture", new FileInfo(imagePath));
                picture.SetSize(70, 50); // Đặt kích thước cho hình ảnh
                picture.From.Row = 1;
                picture.From.Column = 0;

                worksheet.Column(1).Width = 8;
                // Đặt văn bản vào ô A2
                worksheet.Cells["B1"].Value = "CTY CỔ PHẦN DƯỢC PHẨM OPC";
                worksheet.Cells["B1"].Style.Font.Bold = true;
                var cellB1 = worksheet.Cells["B1"];
                cellB1.Style.Font.Bold = true;
                worksheet.Cells["B2"].Value = Dvcs;
                worksheet.Cells["B3"].Value = $"Số: {So}/KT-{Dvcs1}";
                worksheet.Cells["K1"].Value = "Cộng Hòa Xã Hội Chủ Nghĩa Việt Nam";
                worksheet.Cells["K2"].Value = "Độc Lập - Tự Do - Hạnh Phúc";
                worksheet.Cells["K2"].Style.Indent = 4;
                worksheet.Cells["K2"].Style.Font.UnderLine = true;
                worksheet.Cells["E4"].Value = "BẢNG ĐỐI CHIẾU CÔNG NỢ";
                worksheet.Cells["E4"].Style.Font.Bold = true;
                worksheet.Cells["E4"].Style.Font.Size = 16;



                //worksheet.Cells["F4"].Value = $"hàng chưa thanh toán cho chúng tôi tính đến ngày {denngay} là: {tongno}";
                worksheet.Cells["F5"].Value = $"{Time}";

                worksheet.Cells["A6"].Value = $"Tên khách hàng: {ten_dt}";
                worksheet.Cells["A7"].Value = $"I.Số dư nợ trước ngày {TruocNgay}/ {TruocThang}/ {TruocNam} mang sang: {NoDauKy} đồng.";
                worksheet.Cells["A8"].Value = $"II.Doanh thu mua hàng và thanh toán trong kỳ đối chiếu này:";
                worksheet.Cells["A9"].Value = $"1.Doanh thu khách hàng mua trong kỳ: {TienHD}";
                worksheet.Cells["A10"].Value = $"2.Khách hàng đã thanh toán/ trả hàng/ bù trừ trong kỳ: {TienTT} đồng.";
                worksheet.Cells["A11"].Value = $"III.Số tiền khách hàng chưa thanh toán, tính đến cuối ngày {DenNgay}/ {DenThang}/ {DenNam} là: {TonNo} đồng.";
                worksheet.Cells["A12"].Value = $"Số tiền nợ bằng chữ là: {TienChu}";
                worksheet.Cells["A13"].Value = $"Chi tiết các hóa đơn chưa thanh toán: ";
                var startRow = 14;
                var startColumn = 1;

                //var nextII = startRow + 1;
                //var startRowIII = nextII + 3;

                var sttIIICell = worksheet.Cells[startRow + 1, startColumn, startRow + 2, startColumn + 1];
                sttIIICell.Merge = true;
                worksheet.Column(startColumn).Width = 15; // Đặt chiều rộng của cột chứa ô "STT" thành 15 đơn vị.

                sttIIICell.Value = "STT";
                sttIIICell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
                sttIIICell.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc
                sttIIICell.Style.Font.Bold = true;
                sttIIICell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                var HD = worksheet.Cells[startRow + 1, startColumn + 2, startRow + 1, startColumn + 5];
                HD.Merge = true;
                HD.Value = "HÓA ĐƠN";
                HD.Style.Font.Bold = true;
                HD.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                HD.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                HD.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                var SO = worksheet.Cells[startRow + 2, startColumn + 2, startRow + 2, startColumn + 3];
                SO.Merge = true;
                SO.Value = "SỐ";
                worksheet.Column(startColumn + 1).Width = 15;
                SO.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                SO.Style.Font.Bold = true;
                SO.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                SO.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                var NGAY = worksheet.Cells[startRow + 2, startColumn + 4, startRow + 2, startColumn + 5];
                NGAY.Merge = true;
                NGAY.Value = "NGÀY";
                worksheet.Column(startColumn + 2).Width = 15;
                NGAY.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                NGAY.Style.Font.Bold = true;
                NGAY.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                NGAY.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                var TIENHD = worksheet.Cells[startRow + 1, startColumn + 6, startRow + 2, startColumn + 6];
                TIENHD.Merge = true;
                worksheet.Column(startColumn + 6).Width = 30;
                TIENHD.Value = "TIỀN HÓA ĐƠN (ĐÃ TRỪ CODE KM)";
                worksheet.Column(startColumn + 3).Width = 30;
                TIENHD.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                TIENHD.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                TIENHD.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                TIENHD.Style.Font.Bold = true;

                var CODEKM = worksheet.Cells[startRow + 1, startColumn + 7, startRow + 2, startColumn + 8];
                CODEKM.Merge = true;
                CODEKM.Value = "CODE KM";
                worksheet.Column(startColumn + 4).Width = 15;
                worksheet.Column(startColumn + 5).Width = 15;
                worksheet.Column(startColumn + 6).Width = 15;
                worksheet.Column(startColumn + 8).Width = 15;
                CODEKM.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                CODEKM.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                CODEKM.Style.Font.Bold = true;
                CODEKM.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                var GHICHU = worksheet.Cells[startRow + 1, startColumn + 9, startRow + 2, startColumn + 10];
                GHICHU.Merge = true;
                GHICHU.Value = "GHI CHÚ";
                worksheet.Column(startColumn + 4).Width = 15;
                worksheet.Column(startColumn + 5).Width = 15;
                worksheet.Column(startColumn + 6).Width = 15;
                worksheet.Column(startColumn + 8).Width = 15;
                GHICHU.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                GHICHU.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                GHICHU.Style.Font.Bold = true;
                GHICHU.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //// Đảm bảo rằng có dữ liệu trong biến tableData
                if (combinedData != null && combinedData.Any())
                {
                    var stt = 1;
                    // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
                    for (int row = 0; row < combinedData.Count; row++)
                    {
                        var rowData = combinedData[row];
                        var sttCell = worksheet.Cells[startRow + 3 + row, startColumn, startRow + 3 + row, startColumn + 1];
                        sttCell.Merge = true;
                        sttCell.Value = stt;
                        FormatCellNoQH(sttCell);

                        var soCell = worksheet.Cells[startRow + 3 + row, startColumn + 2, startRow + 3 + row, startColumn + 3];
                        soCell.Merge = true;
                        soCell.Value = rowData.So;
                        FormatCellNoQH(soCell);

                        var ngayCell = worksheet.Cells[startRow + 3 + row, startColumn + 4, startRow + 3 + row, startColumn + 5];
                        ngayCell.Merge = true;
                        ngayCell.Value = rowData.Ngay;
                        FormatCellNoQH(ngayCell);

                        var tienhdCell = worksheet.Cells[startRow + 3 + row, startColumn + 6];
                        //tienhdCell.Merge = true;
                        tienhdCell.Value = rowData.TienHD;
                        worksheet.Column(startColumn + 6).Width = 30;
                        FormatCellNoQH(tienhdCell);
                        tienhdCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        var codekm = worksheet.Cells[startRow + 3 + row, startColumn + 7, startRow + 3 + row, startColumn + 8];
                        codekm.Merge = true;

                        FormatCellNoQH(codekm);

                        var ghichu = worksheet.Cells[startRow + 3 + row, startColumn + 9, startRow + 3 + row, startColumn + 10];
                        ghichu.Merge = true;
                        ghichu.Value = rowData.GhiChu;
                        FormatCellNoQH(ghichu);
                        //for (int col = 0; col < rowData.Count; col++)
                        //{
                        //    var cell = worksheet.Cells[startRow + 1 + row, startColumn + col * 2];
                        //    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        //    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //    cell.Value = rowData[col];

                        //    var mergeCell = worksheet.Cells[startRow + 1 + row, startColumn + col * 2, startRow + 1 + row, startColumn + col * 2 + 1];
                        //    mergeCell.Merge = true;
                        //    mergeCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        //}
                        stt++;
                    }
                }
                else
                {
                    worksheet.Cells[startRow, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
                }

                var TC2 = worksheet.Cells[startRow + combinedData.Count + 3, startColumn, startRow + combinedData.Count + 3, startColumn + 5];
                TC2.Merge = true;
                TC2.Value = "Tổng cộng: ";
                TC2.Style.Font.Bold = true;
                TC2.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                TC2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
                TC2.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc

                var TONNO2 = worksheet.Cells[startRow + combinedData.Count + 3, startColumn + 6];

                TONNO2.Value = $"{TonNo2}";
                //worksheet.Cells[startRowIII + 1 + combinedDataHD.Count, startColumn + 6].Value = $"{TonNo2}";
                TONNO2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                TONNO2.Style.Font.Bold = true;
                TONNO2.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count + 3, startColumn + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                var Null1 = worksheet.Cells[startRow + combinedData.Count + 3, startColumn + 7, startRow + combinedData.Count + 3, startColumn + 8];
                Null1.Merge = true;
                Null1.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                var Null2 = worksheet.Cells[startRow + combinedData.Count + 3, startColumn + 9, startRow + combinedData.Count + 3, startColumn + 10];
                Null2.Merge = true;
                Null2.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count + 3, startColumn + 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                var end = startRow + combinedData.Count + 4;
                worksheet.Cells[end + 1, startColumn].Value = $"Quý khách vui lòng xác nhận số nợ trên tại thời điểm ngày {DenNgay}/ {DenThang}/ {DenNam} và gửi về Công Ty Cổ Phần Dược Phẩm OPC - {Dvcs} trước ngày {NgayTT}";
                worksheet.Cells[end + 2, startColumn].Value = $"Địa chỉ: {DiaChi}";
                worksheet.Cells[end + 3, startColumn].Value = $"Điện thoại: {SDT}";
                worksheet.Cells[end + 4, startColumn].Value = $"Số tiền còn nợ đề nghị Quý khách hàng thanh toán bằng tiền mặt hoặc chuyển khoản vào tài khoản {Dvcs}, số tài khoản: {TK}";
                worksheet.Cells[end + 5, startColumn].Value = $"Nếu có gì không rõ vui lòng liên hệ: {LH}";
                worksheet.Cells[end + 7, startColumn].Value = "Trân trọng cảm ơn!";
                worksheet.Cells[end + 7, startColumn].Style.Font.Bold = true;
                worksheet.Cells[end + 7, startColumn].Style.Font.Italic = true;
                worksheet.Cells[end + 8, startColumn + 9].Value = $"Ngày {NgayKy} tháng {ThangKy} năm {NamKy}";
                worksheet.Cells[end + 8, startColumn + 9].Style.Font.Bold = true;
                worksheet.Cells[end + 8, startColumn + 9].Style.Indent = 2;
                worksheet.Cells[end + 9, startColumn + 1].Value = "KHÁCH HÀNG XÁC NHẬN";
                worksheet.Cells[end + 9, startColumn + 1].Style.Font.Bold = true;

                worksheet.Cells[end + 9, startColumn + 5].Value = "GIÁM ĐỐC";
                worksheet.Cells[end + 9, startColumn + 5].Style.Font.Bold = true;

                worksheet.Cells[end + 9, startColumn + 10].Value = "KẾ TOÁN";
                worksheet.Cells[end + 9, startColumn + 10].Style.Font.Bold = true;

                worksheet.Cells[end + 10, startColumn].Value = "(Ký, đóng dấu, ghi rõ họ tên)";
                worksheet.Cells[end + 10, startColumn].Style.Font.Bold = true;
                worksheet.Cells[end + 10, startColumn].Style.Indent = 5;


                package.Save();
                byte[] fileBytes = package.GetAsByteArray();

                // Trả về tệp Excel dưới dạng dữ liệu binary
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

            }
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
        public ActionResult PhieuNhapKho_Fill()
        {
            List<BKHoaDonGiaoHang> dmDlistVT = LoadDmVt();
            ViewBag.DataVT = dmDlistVT;
            return View();
        }
        public ActionResult PhieuNhapKho_Index(string Ma_Dvcs)
        {
            DataSet ds = new DataSet();
            connectSQL();

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_MauInChungTuNKXK_SAP]";
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            var Dvcs = Request.Cookies["Dvcs3"].Value;
            var LoaiCt = Request.Cookies["LoaiCt"].Value;
            var MaKho = Request.Cookies["Ma_Dv"].Value;
            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Loai_Ct", LoaiCt);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", Dvcs);
                    cmd.Parameters.AddWithValue("@_Ma_Kho", MaKho);
                    sda.Fill(ds);

                }
            }
            return View(ds);

        }
        public ActionResult PhieuNhapKho(string SoCt)
        {
            DataSet ds = new DataSet();
            connectSQL();
            var So_Ct = Request.Cookies["So_Ct"].Value;
            //var So_Ct = SoCt;
            var LoaiCt = Request.Cookies["LoaiCt"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_MauInChungTuNKXK_SAP]";



            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                var fromDate = Request.Cookies["From_date"].Value;
                var toDate = Request.Cookies["To_Date"].Value;
                var Dvcs = Request.Cookies["Dvcs3"].Value;

                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Loai_Ct", LoaiCt);
                    cmd.Parameters.AddWithValue("@_So_Ct", So_Ct);



                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult TonNo()
        {
            List<MauInChungTu> dmDlist = LoadDmDt("");
            List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            ViewBag.DataItems = dmDlist;
            ViewBag.DataTDV = dmDlistTDV;
            DataSet ds = new DataSet();
            connectSQL();

            //var So_Ct = SoCt;
            var ma_Cbnv = Request.Cookies["Ma_TDV"].Value;
            var ma_dvcs = Request.Cookies["MA_DVCS"].Value;
            var ma_dt = Request.Cookies["Ma_Dt"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_RpCongNoChiTietTDV_SAP]";



            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                var fromDate = Request.Cookies["From_date"].Value;
                var toDate = Request.Cookies["To_Date"].Value;


                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Ma_DvCs", ma_dvcs);
                    cmd.Parameters.AddWithValue("@_Ma_CbNv", ma_Cbnv);
                    cmd.Parameters.AddWithValue("@_Ma_Dt", ma_dt);


                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public List<GetData> LoadData()
        {
            connectSQL();
            List<GetData> dataItems = new List<GetData>();
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_RpCongNoChiTietTDV_SAP]", connection))
                {
                    var fromDate = Request.Cookies["From_date"].Value;
                    var toDate = Request.Cookies["To_Date"].Value;
                    var ma_Cbnv = Request.Cookies["Ma_TDV"].Value;
                    var ma_dvcs = Request.Cookies["MA_DVCS"].Value;
                    var ma_dt = Request.Cookies["Ma_Dt"].Value;
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;
                    using (SqlDataAdapter sda = new SqlDataAdapter(command))
                    {
                        DataSet ds = new DataSet();
                        command.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                        command.Parameters.AddWithValue("@_Den_Ngay", toDate);
                        command.Parameters.AddWithValue("@_Ma_DvCs", ma_dvcs);
                        command.Parameters.AddWithValue("@_Ma_CbNv", ma_Cbnv);
                        command.Parameters.AddWithValue("@_Ma_Dt", ma_dt);
                        sda.Fill(ds);

                        // Kiểm tra xem DataSet có bảng dữ liệu hay không
                        if (ds.Tables.Count > 0)
                        {
                            DataTable dt = ds.Tables[0];

                            foreach (DataRow row in dt.Rows)
                            {
                                GetData dataItem = new GetData
                                {
                                    TenDt = row["Ten_Dt"].ToString(),
                                    NgayCt = row["Ngay_Ct1"].ToString(),
                                    SoCtEinv = row["So_Ct_Einv"].ToString(),
                                    NgayDenHan = row["Ngay_Den_Han1"].ToString(),
                                    CongNoTT = Convert.ToDecimal(row["Cong_No_TT"].ToString()),
                                    TienThue = Convert.ToDecimal(row["Tien_Thue"].ToString()),
                                    CongNo = Convert.ToDecimal(row["Cong_No"].ToString()),
                                    //TotalCongNoTT = Convert.ToDecimal(row["Tong_CN_TT"].ToString()),
                                    //TotalCongNoST = Convert.ToDecimal(row["Tong_CN_ST"].ToString()),
                                    //TotalCongNo = Convert.ToDecimal(row["Tong_CN"].ToString()),



                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }
            return dataItems;
        }
        public ActionResult TonNo_Fill()
        {
            //List<MauInChungTu> dmDlist = LoadDmDt("");
            //List<BKHoaDonGiaoHang> dmDListDt = LoadDmDt3();
            //ViewBag.DataItems = dmDlist;
            DataSet ds = new DataSet();
            connectSQL();

            //var So_Ct = SoCt;
            var ma_Cbnv = Request.Cookies["Ma_TDV"].Value;
            var ma_dvcs = Request.Cookies["MA_DVCS"].Value;
            var ma_dt = Request.Cookies["Ma_Dt"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_RpCongNoChiTietTDV_SAP]";



            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                var fromDate = Request.Cookies["From_date"].Value;
                var toDate = Request.Cookies["To_Date"].Value;


                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Ma_DvCs", ma_dvcs);
                    cmd.Parameters.AddWithValue("@_Ma_CbNv", ma_Cbnv);
                    cmd.Parameters.AddWithValue("@_Ma_Dt", ma_dt);


                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult TonNo_Index()
        {
            string ma_dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : string.Empty;
            List<MauInChungTu> dmDlist = LoadDmDt("");
            List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            ViewBag.DataItems = dmDlist;
            ViewBag.DataTDV = dmDlistTDV;
            return View();
        }
        public List<BKHoaDonGiaoHang> LoadDmTDV()
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            connectSQL();

            List<BKHoaDonGiaoHang> dataItems = new List<BKHoaDonGiaoHang>();

            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DanhSachTDV]", connection))
                {
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);

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

        private void FormatCell(ExcelRangeBase cell)
        {
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            cell.Style.Font.Bold = true;
            // Các định dạng khác của ô có thể thêm vào tại đây
        }
        private void FormatCellNoQH(ExcelRangeBase cell)
        {
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // Các định dạng khác của ô có thể thêm vào tại đây
        }
        void CreateEmptyRow(ExcelWorksheet worksheet, int startColumn, int currentRow, decimal tongCongNoTT, decimal tongCongNoST, decimal tongCongNo)
        {
            var sttCell = worksheet.Cells[currentRow, startColumn, currentRow, startColumn + 6];
            sttCell.Merge = true;
            sttCell.Value = "Tổng cộng:";
            //worksheet.Cells[currentRow, startColumn].Value = "Tổng cộng";
            FormatCell(sttCell);

            // Các phần khác của mã

            worksheet.Cells[currentRow, startColumn + 7].Value = $"{tongCongNoTT}";
            FormatCell(worksheet.Cells[currentRow, startColumn + 7]);
            worksheet.Cells[currentRow, startColumn + 8].Value = $"{tongCongNoST}";
            FormatCell(worksheet.Cells[currentRow, startColumn + 8]);
            worksheet.Cells[currentRow, startColumn + 9].Value = $"{tongCongNo}";
            FormatCell(worksheet.Cells[currentRow, startColumn + 9]);
        }

        public ActionResult ExportSoTonNo()
        {
            var fileName = $"SoTonNo{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            // Lấy dữ liệu từ cookie

            List<GetData> combinedData = LoadData();
            // Kiểm tra xem có dữ liệu từ cookie không



            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("MySheet");
                worksheet.View.ShowGridLines = false;

                string maTDV = Request.Cookies["Ma_CbNv"] != null ? HttpUtility.UrlDecode(Request.Cookies["Ma_CbNv"].Value) : "";
                string tenTDV = Request.Cookies["Ten_TDV"] != null ? HttpUtility.UrlDecode(Request.Cookies["Ten_TDV"].Value) : "";
                string ngay = Request.Cookies["NgayDaChinhSua"] != null ? HttpUtility.UrlDecode(Request.Cookies["NgayDaChinhSua"].Value) : "";
                string TongCN = Request.Cookies["TongCN"] != null ? HttpUtility.UrlDecode(Request.Cookies["TongCN"].Value) : "";
                string TongCNTT = Request.Cookies["TongCNTT"] != null ? HttpUtility.UrlDecode(Request.Cookies["TongCNTT"].Value) : "";
                string TongCNST = Request.Cookies["TongCNST"] != null ? HttpUtility.UrlDecode(Request.Cookies["TongCNST"].Value) : "";
                worksheet.Cells.Style.Font.Name = "Times New Roman";

                worksheet.Cells["C3"].Value = $"SỔ CHI TIẾT TỒN NỢ PHẢI THU ĐẾN NGÀY {ngay}";
                worksheet.Cells["C3"].Style.Font.Bold = true;
                worksheet.Cells["C3"].Style.Font.Size = 16;

                worksheet.Cells["E4"].Value = $"(Kế hoạch thu từ ngày .../... đến ngày.../...)";
                worksheet.Cells["E4"].Style.Font.Bold = true;
                worksheet.Cells["A6"].Value = $"Trình Dược Viên: {maTDV} - {tenTDV}";
                worksheet.Cells["A7"].Value = "Nhân Viên Giao Nhận: ";
                worksheet.Cells["A8"].Value = $"Kế Toán Công Nợ: ";
                var startRow = 11;
                var startColumn = 1;
                worksheet.Cells[startRow - 1, startColumn].Value = "SỐ";
                worksheet.Cells[startRow - 1, startColumn + 1].Value = "KHÁCH HÀNG";
                worksheet.Cells[startRow - 1, startColumn + 2].Value = "NGÀY HÓA ĐƠN";
                worksheet.Cells[startRow - 1, startColumn + 3].Value = "SỐ HÓA ĐƠN";
                worksheet.Cells[startRow - 1, startColumn + 4].Value = "NGÀY THANH TOÁN";
                worksheet.Cells[startRow - 1, startColumn + 5].Value = "GN THU";
                worksheet.Cells[startRow - 1, startColumn + 6].Value = "TDV THU";
                worksheet.Cells[startRow - 1, startColumn + 7].Value = "TIỀN TRƯỚC THUẾ";
                worksheet.Cells[startRow - 1, startColumn + 8].Value = "TIỀN THUẾ";
                worksheet.Cells[startRow - 1, startColumn + 9].Value = "TỔNG TIỀN";
                for (int col = 0; col < 10; col++)
                {
                    var columnHeaderCell = worksheet.Cells[startRow - 1, startColumn + col];
                    columnHeaderCell.Style.Font.Bold = true;
                    columnHeaderCell.Style.Font.Size = 10;
                    columnHeaderCell.Style.Font.Color.SetColor(Color.Black);
                    columnHeaderCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    columnHeaderCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    columnHeaderCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    columnHeaderCell.Style.Fill.BackgroundColor.SetColor(Color.White);
                    columnHeaderCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                }

                worksheet.Column(startColumn).Width = 10;
                worksheet.Column(startColumn + 1).Width = 35;
                worksheet.Column(startColumn + 2).Width = 25;
                worksheet.Column(startColumn + 3).Width = 15;
                worksheet.Column(startColumn + 4).Width = 25;
                worksheet.Column(startColumn + 5).Width = 15;
                worksheet.Column(startColumn + 6).Width = 25;
                worksheet.Column(startColumn + 7).Width = 25;
                worksheet.Column(startColumn + 8).Width = 25;
                worksheet.Column(startColumn + 9).Width = 25;
                var end = 1;
                if (combinedData != null && combinedData.Any())
                {
                    // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
                    var stt = 1;
                    int currentRow = startRow;
                    string previousTenDt = null;
                    Dictionary<string, int> tongCongNoTT = new Dictionary<string, int>();
                    Dictionary<string, int> tongCongNoST = new Dictionary<string, int>();
                    Dictionary<string, int> tongCongNo = new Dictionary<string, int>();
                    for (int row = 0; row < combinedData.Count; row++)
                    {
                        var rowData = combinedData[row];

                        // Kiểm tra nếu giá trị TenDt thay đổi
                        if (rowData.TenDt != previousTenDt)
                        {
                            // Tạo hàng rỗng nếu không phải là hàng đầu tiên
                            if (currentRow > startRow)
                            {
                                CreateEmptyRow(worksheet, startColumn, currentRow,
                                 tongCongNoTT[previousTenDt],
                                 tongCongNoST[previousTenDt],
                                 tongCongNo[previousTenDt]);
                                currentRow++;
                            }

                            // Ghi tên công ty vào ô đầu tiên
                            worksheet.Cells[currentRow, startColumn + 1].Value = rowData.TenDt;
                            var cell = worksheet.Cells[currentRow, startColumn + 1];
                            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; // Đặt giá trị canh lề
                            FormatCell(cell);


                            // Cập nhật giá trị previousTenDt
                            previousTenDt = rowData.TenDt;

                            // Reset số thứ tự
                            stt = 1;
                            if (!tongCongNoTT.ContainsKey(rowData.TenDt))
                            {
                                tongCongNoTT[rowData.TenDt] = 0;
                                tongCongNoST[rowData.TenDt] = 0;
                                tongCongNo[rowData.TenDt] = 0;
                            }
                        }
                        else
                        {
                            // Nếu giá trị TenDt trùng nhau, bỏ trống giá trị TenDt
                            worksheet.Cells[currentRow, startColumn + 1].Value = "";
                            FormatCell(worksheet.Cells[currentRow, startColumn + 1]);
                        }

                        // Ghi các giá trị cột khác
                        worksheet.Cells[currentRow, startColumn].Value = stt;
                        FormatCell(worksheet.Cells[currentRow, startColumn]);
                        worksheet.Cells[currentRow, startColumn + 2].Value = rowData.NgayCt;
                        FormatCell(worksheet.Cells[currentRow, startColumn + 2]);
                        worksheet.Cells[currentRow, startColumn + 3].Value = rowData.SoCtEinv;
                        FormatCell(worksheet.Cells[currentRow, startColumn + 3]);
                        worksheet.Cells[currentRow, startColumn + 4].Value = rowData.NgayDenHan;
                        FormatCell(worksheet.Cells[currentRow, startColumn + 4]);
                        FormatCell(worksheet.Cells[currentRow, startColumn + 5]);
                        worksheet.Cells[currentRow, startColumn + 6].Value = rowData.GhiChu;
                        FormatCell(worksheet.Cells[currentRow, startColumn + 6]);
                        worksheet.Cells[currentRow, startColumn + 7].Value = rowData.CongNoTT;
                        FormatCell(worksheet.Cells[currentRow, startColumn + 7]);
                        worksheet.Cells[currentRow, startColumn + 8].Value = rowData.TienThue;
                        FormatCell(worksheet.Cells[currentRow, startColumn + 8]);
                        worksheet.Cells[currentRow, startColumn + 9].Value = rowData.CongNo;
                        FormatCell(worksheet.Cells[currentRow, startColumn + 9]);

                        // Tăng số thứ tự và dòng
                        stt++;
                        currentRow++;
                        tongCongNoTT[rowData.TenDt] += (int)rowData.CongNoTT;
                        tongCongNoST[rowData.TenDt] += (int)rowData.TienThue;
                        tongCongNo[rowData.TenDt] += (int)rowData.CongNo;
                    }
                    CreateEmptyRow(worksheet, startColumn, currentRow,
                  tongCongNoTT[previousTenDt],
                  tongCongNoST[previousTenDt],
                  tongCongNo[previousTenDt]);
                    var sttCell = worksheet.Cells[currentRow + 1, startColumn, currentRow + 1, startColumn + 6];
                    sttCell.Merge = true;
                    sttCell.Value = "TỔNG CÔNG NỢ";

                    FormatCell(sttCell);
                    FormatCell(worksheet.Cells[currentRow + 1, startColumn]);
                    FormatCell(worksheet.Cells[currentRow + 1, startColumn + 2]);
                    FormatCell(worksheet.Cells[currentRow + 1, startColumn + 3]);
                    FormatCell(worksheet.Cells[currentRow + 1, startColumn + 4]);
                    FormatCell(worksheet.Cells[currentRow + 1, startColumn + 5]);
                    FormatCell(worksheet.Cells[currentRow + 1, startColumn + 6]);
                    worksheet.Cells[currentRow + 1, startColumn + 7].Value = $"{TongCNTT}";
                    FormatCell(worksheet.Cells[currentRow + 1, startColumn + 7]);
                    worksheet.Cells[currentRow + 1, startColumn + 8].Value = $"{TongCNST}";
                    FormatCell(worksheet.Cells[currentRow + 1, startColumn + 8]);
                    worksheet.Cells[currentRow + 1, startColumn + 9].Value = $"{TongCN}";
                    FormatCell(worksheet.Cells[currentRow + 1, startColumn + 9]);
                    end = currentRow + 1;
                }
                else
                {
                    worksheet.Cells[startRow, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
                    FormatCell(worksheet.Cells[startRow, startColumn]);
                }
                //worksheet.Cells[startRow + combinedData.Count -2, startColumn + 1].Value = "Tổng cộng";
                //worksheet.Cells[startRow + combinedData.Count-2, startColumn + 1].Style.Font.Bold = true;
                //worksheet.Cells[startRow + combinedData.Count-2, startColumn].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                //worksheet.Cells[startRow + combinedData.Count-2, startColumn + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                //worksheet.Cells[startRow + combinedData.Count-2, startColumn + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                //worksheet.Cells[startRow + combinedData.Count-2, startColumn + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                //worksheet.Cells[startRow + combinedData.Count - 2, startColumn + 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                //worksheet.Cells[startRow + combinedData.Count - 2, startColumn + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                //worksheet.Cells[startRow + combinedData.Count - 2, startColumn + 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                //worksheet.Cells[startRow + combinedData.Count - 2, startColumn + 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                var endRow = startRow + end + 1 - 11;

                worksheet.Cells[endRow, startColumn + 1].Value = "BÁO CÁO CỦA NGƯỜI PHỤ TRÁCH THU: ";
                worksheet.Cells[endRow + 1, startColumn + 1].Value = "-Tổng giá trị nợ:.................................................................. ";
                worksheet.Cells[endRow + 2, startColumn + 1].Value = "-Tổng giá trị kế hoạch giao nhận thu:.................................. ";
                worksheet.Cells[endRow + 3, startColumn + 1].Value = "-Tổng giá trị nợ giao nhận thu được:................................... ";
                worksheet.Cells[endRow + 4, startColumn + 1].Value = "-Tổng giá trị kế hoạch khách hàng CK:............................... ";
                worksheet.Cells[endRow + 5, startColumn + 1].Value = "-Tổng giá trị nợ khách hàng CK:.........................................";
                worksheet.Cells[endRow + 6, startColumn + 1].Value = "-Tổng giá trị kế hoạch trình dược viên thu:......................... ";
                worksheet.Cells[endRow + 7, startColumn + 1].Value = "-Tổng giá trị nợ trình dược viên thu được:.......................... ";



                worksheet.Cells[endRow + 1, startColumn + 4].Value = "-Tỷ lệ % kế hoạch GN thu so với nợ:................................";
                worksheet.Cells[endRow + 2, startColumn + 4].Value = "-Tỷ lệ % GN thu được so với kế hoạch:............................";
                worksheet.Cells[endRow + 3, startColumn + 4].Value = "-Tỷ lệ % KH CK so với nợ:............................................... ";
                worksheet.Cells[endRow + 4, startColumn + 4].Value = "-Tỷ lệ % KH CK so với kế hoạch: ....................................";
                worksheet.Cells[endRow + 5, startColumn + 4].Value = "-Tỷ lệ % kế hoạch TDV thu so với nợ:..............................";
                worksheet.Cells[endRow + 6, startColumn + 4].Value = "-Tỷ lệ % TDV thu được so với kế hoạch: ..........................";
                int nextRow = endRow + 7;
                worksheet.Cells[nextRow, startColumn + 7].Value = $"Ngày     tháng     năm";
                worksheet.Cells[nextRow, startColumn + 7].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 1, startColumn + 1].Value = $"PGĐ.Tài Chính";
                worksheet.Cells[nextRow + 1, startColumn + 1].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 1, startColumn + 2].Value = $"Giám Sát";
                worksheet.Cells[nextRow + 1, startColumn + 2].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 1, startColumn + 2].Style.Indent = 3;
                worksheet.Cells[nextRow + 2, startColumn + 2].Value = "(Ký, họ tên)";
                worksheet.Cells[nextRow + 2, startColumn + 2].Style.Font.Italic = true;
                worksheet.Cells[nextRow + 2, startColumn + 2].Style.Indent = 3;
                worksheet.Cells[nextRow + 2, startColumn + 2].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 1, startColumn + 1].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 1, startColumn + 4].Value = $"Kế Toán Công Nợ";
                worksheet.Cells[nextRow + 2, startColumn + 4].Value = "(Ký, họ tên)";
                worksheet.Cells[nextRow + 2, startColumn + 4].Style.Font.Italic = true;
                worksheet.Cells[nextRow + 2, startColumn + 4].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 2, startColumn + 4].Style.Indent = 7;
                worksheet.Cells[nextRow + 1, startColumn + 4].Style.Indent = 7;
                worksheet.Cells[nextRow + 1, startColumn + 4].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 2, startColumn + 1].Value = "(Ký, họ tên)";
                worksheet.Cells[nextRow + 2, startColumn + 1].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 2, startColumn + 1].Style.Font.Italic = true;
                worksheet.Cells[nextRow + 1, startColumn + 7].Value = $"Người Phụ Trách Thu";
                worksheet.Cells[nextRow + 2, startColumn + 7].Value = "(Ký, họ tên)";
                worksheet.Cells[nextRow + 1, startColumn + 7].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 2, startColumn + 7].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 2, startColumn + 7].Style.Font.Italic = true;


                package.Save();
                byte[] fileBytes = package.GetAsByteArray();

                // Trả về tệp Excel dưới dạng dữ liệu binary
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

            }




        }
        public ActionResult BienBanBanGiaoNTHH()
        {
            DataSet ds = new DataSet();
            connectSQL();

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_DanhSachHoaDonBBBG_SAP]";
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;


            var Dvcs = Request.Cookies["MA_DVCS"].Value == "" ? Request.Cookies["Dvcs3"].Value : Request.Cookies["MA_DVCS"].Value;
            //var MaTDV = Request.Cookies["Ma_CbNv"] != null ? Request.Cookies["Ma_CbNv"].Value : string.Empty;
            var MaDt = Request.Cookies["Ma_Dt"] != null ? Request.Cookies["Ma_Dt"].Value : string.Empty;
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

                    cmd.Parameters.AddWithValue("@_ma_dvcs", Dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public List<BKHoaDonGiaoHang> LoadDmDt3()
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"] != null ? Request.Cookies["Ma_dvcs"].Value : string.Empty;
            connectSQL();

            List<BKHoaDonGiaoHang> dataItems = new List<BKHoaDonGiaoHang>();
            string appendedString = ma_dvcs == "OPC_B1" ? "_010203" : "_01"; // Dòng này cộng chuỗi dựa trên giá trị của Ma_dvcs
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DmDtTdv_SAP_MauIn]", connection))
                {
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.AddWithValue("@_Ma_Dvcs", ma_dvcs + appendedString);

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
                                    Ma_CbNv = row["Ma_Dt"].ToString(),
                                    hoten = row["Ten_Dt"].ToString(),
                                    Ma_Dvcs = row["Dvcs"].ToString()
                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }

            return dataItems;
        }
        public List<BKHoaDonGiaoHang> LoadDmDt1()
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"] != null ? Request.Cookies["Ma_dvcs"].Value : string.Empty;
            connectSQL();

            List<BKHoaDonGiaoHang> dataItems = new List<BKHoaDonGiaoHang>();
            string appendedString = ma_dvcs == "OPC_B1" ? "_010203" : "_01"; // Dòng này cộng chuỗi dựa trên giá trị của Ma_dvcs
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DmDtTdv_SAP_MauIn]", connection))
                {
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.AddWithValue("@_Ma_Dvcs", ma_dvcs + appendedString);

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
                                    Ma_CbNv = row["Ma_Dt"].ToString(),
                                    hoten = row["Ten_Dt"].ToString(),
                                    Ma_Dvcs = row["Dvcs"].ToString()
                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }

            return dataItems;
        }
        public ActionResult BienBanBanGiaoNTHH_Index(MauInChungTu MauIn)
        {
            DataSet ds = new DataSet();
            connectSQL();
            List<MauInChungTu> dmDlist = LoadDmDt("");

            ViewBag.DataItems = dmDlist;

            string Pname = "[usp_DanhSachHoaDonBBBG_SAP]";
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;

            var Dvcs = Request.Cookies["MA_DVCS"].Value;

            var MaDt = Request.Cookies["Ma_Dt"] != null ? Request.Cookies["Ma_Dt"].Value : string.Empty;


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
                    //cmd.Parameters.AddWithValue("@_Ma_CbNv", MaTDV);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", Dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult BienBanBanGiaoNTHH_Fill()
        {
            List<MauInChungTu> dmDlist = LoadDmDt("");

            ViewBag.DataItems = dmDlist;
            return View();
        }
        public ActionResult PhieuXacNhanTTTCK_Fill()
        {
            string ma_dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : string.Empty;
            List<MauInChungTu> dmDlist = LoadDmDt("");
            List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            ViewBag.DataItems = dmDlist;
            ViewBag.DataTDV = dmDlistTDV;
            return View();
        }
        public ActionResult PhieuXacNhanTTTCK2_Fill()
        {
            string ma_dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : string.Empty;
            List<MauInChungTu> dmDlist = LoadDmDt("");
            List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            ViewBag.DataItems = dmDlist;
            ViewBag.DataTDV = dmDlistTDV;
            return View();
        }
        public ActionResult PhieuXacNhanTTTCK_In()
        {
            List<MauInChungTu> dmDlist = LoadDmDt("");
            List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            string ma_dvcs = Request.Cookies["MA_DVCS"].Value;
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            var MaDt = Request.Cookies["Ma_Dt"] != null ? Request.Cookies["Ma_Dt"].Value : string.Empty;
            var MaTDV = Request.Cookies["Ma_TDV"].Value;

            DataSet ds = new DataSet();

            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataItems = dmDlist;
            connectSQL();
            //var SoCT = Request.Cookies["So_Ct"] != null ? Request.Cookies["So_Ct"].Value : "";
            //MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_XacNhanThanhToanCKTT_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                //MauIn.From_date = Request.Cookies["From_date"].Value;
                //MauIn.To_date = Request.Cookies["To_Date"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Ma_Dt", MaDt);

                    cmd.Parameters.AddWithValue("@_Ma_CbNv", MaTDV);
                    cmd.Parameters.AddWithValue("@_Ma_DvCs", ma_dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult PhieuXacNhanTTTCK2_In()
        {
            List<MauInChungTu> dmDlist = LoadDmDt("");
            List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            string ma_dvcs = Request.Cookies["MA_DVCS"].Value;
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            var MaDt = Request.Cookies["Ma_Dt"] != null ? Request.Cookies["Ma_Dt"].Value : string.Empty;
            var MaTDV = Request.Cookies["Ma_TDV"].Value;

            DataSet ds = new DataSet();

            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataItems = dmDlist;
            connectSQL();
            //var SoCT = Request.Cookies["So_Ct"] != null ? Request.Cookies["So_Ct"].Value : "";
            //MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_XacNhanThanhToanCKTT_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                //MauIn.From_date = Request.Cookies["From_date"].Value;
                //MauIn.To_date = Request.Cookies["To_Date"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Ma_Dt", MaDt);

                    cmd.Parameters.AddWithValue("@_Ma_CbNv", MaTDV);
                    cmd.Parameters.AddWithValue("@_Ma_DvCs", ma_dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult PhieuXacNhanTTTCK()
        {
            List<MauInChungTu> dmDlist = LoadDmDt("");
            List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            string ma_dvcs = Request.Cookies["MA_DVCS"].Value;
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            var MaDt = Request.Cookies["Ma_Dt"] != null ? Request.Cookies["Ma_Dt"].Value : string.Empty;
            var MaTDV = Request.Cookies["Ma_TDV"].Value;

            DataSet ds = new DataSet();

            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataItems = dmDlist;
            connectSQL();
            //var SoCT = Request.Cookies["So_Ct"] != null ? Request.Cookies["So_Ct"].Value : "";
            //MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_XacNhanThanhToanCKTT_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                //MauIn.From_date = Request.Cookies["From_date"].Value;
                //MauIn.To_date = Request.Cookies["To_Date"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Ma_Dt", MaDt);

                    cmd.Parameters.AddWithValue("@_Ma_CbNv", MaTDV);
                    cmd.Parameters.AddWithValue("@_Ma_DvCs", ma_dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult PhieuXacNhanTTTCK2()
        {
            List<MauInChungTu> dmDlist = LoadDmDt("");
            List<BKHoaDonGiaoHang> dmDlistTDV = LoadDmTDV();
            string ma_dvcs = Request.Cookies["MA_DVCS"].Value;
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            var MaDt = Request.Cookies["Ma_Dt"] != null ? Request.Cookies["Ma_Dt"].Value : string.Empty;
            var MaTDV = Request.Cookies["Ma_TDV"].Value;

            DataSet ds = new DataSet();

            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataItems = dmDlist;
            connectSQL();
            //var SoCT = Request.Cookies["So_Ct"] != null ? Request.Cookies["So_Ct"].Value : "";
            //MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_XacNhanThanhToanCKTT_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                //MauIn.From_date = Request.Cookies["From_date"].Value;
                //MauIn.To_date = Request.Cookies["To_Date"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_Ma_Dt", MaDt);

                    cmd.Parameters.AddWithValue("@_Ma_CbNv", MaTDV);
                    cmd.Parameters.AddWithValue("@_Ma_DvCs", ma_dvcs);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
    }

}