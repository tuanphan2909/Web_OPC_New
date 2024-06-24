using ClosedXML.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using web4.Models;
using System.Drawing;
namespace web4.Controllers
{
    public class TheoDoiGiaoHangController : Controller
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

                using (SqlCommand command = new SqlCommand("[usp_DanhSachHoaDonGiaoHang_SAP]", connection))
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
                                    So_HD = row["so_ct"].ToString(),
                                    Ngay_HD = row["Ngay_Ct1"].ToString(),
                                    Ma_Dt = row["Ma_dt"].ToString(),
                                    Ten_Dt = row["Ten_Dt"].ToString(),
                                    Ma_NVGH = row["Ma_nvgh"].ToString(),
                                    Tien_HD = float.Parse(row["tien"].ToString()),
                                    Tien_Phai_Thu = float.Parse(row["Tien_Phai_Thu"].ToString())


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

                using (SqlCommand command = new SqlCommand("[usp_DanhSachHoaDonGiaoHang_SAP]", connection))
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
                                    So_HD = row["so_ct"].ToString(),
                                    Ngay_HD = row["Ngay_Ct1"].ToString(),
                                    Ma_Dt = row["Ma_dt"].ToString(),
                                    Ten_Dt = row["Ten_Dt"].ToString(),
                                    Ma_NVGH = row["Ma_nvgh"].ToString(),
                                    Tien_HD = float.Parse(row["tien"].ToString()),

                                    Tien_Phai_Thu = float.Parse(row["Tien_Phai_Thu"].ToString())

                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }

            return dataItems;
        }

        public ActionResult InsertGiaoHang()
        {
            List<TheoDoiGiaoHang> dmDlistTDV = LoadDmTDV();

            ViewBag.DataTDV = dmDlistTDV;

            return View();
        }
        public ActionResult InsetGiaoHangLoadHD()
        {
            List<TheoDoiGiaoHang> dmDlistTDV = LoadDmTDV();
            List<TheoDoiGiaoHang> dmListHD = LoadHD();
            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataHD = dmListHD;
            return View();
        }

        public ActionResult UpdateGiaoHangLoadHD()
        {
            List<TheoDoiGiaoHang> dmDlistTDV = LoadDmTDV();
            List<TheoDoiGiaoHang> dmListHD = LoadHD();
            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataHD = dmListHD;
            return View();
        }
        public ActionResult SaveHD(TheoDoiGiaoHang TDGH)
        {
            TDGH.Dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : "";
            TDGH.Ma_NVGH = Request.Cookies["Ma_NVGH"] != null ? Request.Cookies["Ma_NVGH"].Value : "";
            TDGH.Ten_NVGH = Request.Cookies["Ten_NVGH"] != null ? Request.Cookies["Ten_NVGH"].Value : "";
            TDGH.NV_GiaoNhan = Request.Cookies["NV_GiaoNhan"] != null ? Request.Cookies["NV_GiaoNhan"].Value : "";
            TDGH.Ly_do = Request.Cookies["Ly_Do"] != null ? Request.Cookies["Ly_Do"].Value : "";


            string result = "Error!";
            connectSQL();
            if (TDGH != null && TDGH.Details != null)
            {
                try
                {
                    var detailsTable = new DataTable();
                    detailsTable.Columns.Add("So_Hd", typeof(int));
                    detailsTable.Columns.Add("Ngay_HD", typeof(string));
                    detailsTable.Columns.Add("Ma_Dt", typeof(int));
                    detailsTable.Columns.Add("Ten_Dt", typeof(string));
                    detailsTable.Columns.Add("NV_GN", typeof(string));
                    detailsTable.Columns.Add("Giao_HD", typeof(bool));
                    detailsTable.Columns.Add("Tien", typeof(float));
                    detailsTable.Columns.Add("Tien_Phai_Thu", typeof(float));
                    detailsTable.Columns.Add("Noi_Dung", typeof(string));
                    detailsTable.Columns.Add("Chua_giao_hang", typeof(bool));
                    foreach (var detail in TDGH.Details)
                    {
                        detailsTable.Rows.Add(detail.So_Hd, detail.Ngay_HD, detail.Ma_Dt, detail.Ten_Dt, detail.NV_GiaoNhan, detail.Giao_HD, detail.Tien_HD,detail.Tien_Phai_Thu, detail.Noi_Dung, detail.Chua_giao_hang);
                    }

                    using (var connection = new SqlConnection(con.ConnectionString))
                    {
                        connection.Open();

                        using (var command = new SqlCommand("InsertTheoDoiGiaoHang_SAP", connection))
                        {
                            command.CommandType = CommandType.StoredProcedure;

                            command.Parameters.AddWithValue("@_Ngay_Ct", TDGH.Ngay_Ct);
                            command.Parameters.AddWithValue("@_so_Ct", TDGH.So_Ct);
                            command.Parameters.AddWithValue("@_NV_GiaoHang", TDGH.Ma_NVGH);
                            command.Parameters.AddWithValue("@_Ten_NVGiaoHang", TDGH.Ten_NVGH);
                            command.Parameters.AddWithValue("@_Ten_NVPhuKho", TDGH.NV_GiaoNhan);
                            command.Parameters.AddWithValue("@_Dvcs", TDGH.Dvcs);
                            command.Parameters.AddWithValue("@_Ly_Do", TDGH.Ly_do);

                            // Pass details as a TVP parameter
                            var detailsParam = command.Parameters.AddWithValue("@_Details", detailsTable);
                            detailsParam.SqlDbType = SqlDbType.Structured;
                            detailsParam.TypeName = "B30GdetailType"; // Replace with your actual type name

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
            string Pname = "DanhSachTheoDoiGiaoHang";


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
        public ActionResult MauInGiaoHang()
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
        public ActionResult UpdateGiaoHang()
        {
            List<TheoDoiGiaoHang> dmDlistTDV = LoadDmTDV();
            ///List<TheoDoiGiaoHang> dmListHD = LoadHD();
            ViewBag.DataTDV = dmDlistTDV;
            //ViewBag.DataHD = dmListHD;

            DataSet ds = new DataSet();
            connectSQL();

            string Pname = "[EditTheoDoiGiaoHang]";
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
        public ActionResult SaveUpdate(TheoDoiGiaoHang TDGH)
        {



            TDGH.Dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : "";
            TDGH.Ma_NVGH = Request.Cookies["Ma_NVGH"] != null ? Request.Cookies["Ma_NVGH"].Value : "";
            TDGH.Ten_NVGH = Request.Cookies["Ten_NVGH"] != null ? Request.Cookies["Ten_NVGH"].Value : "";
            TDGH.NV_GiaoNhan = Request.Cookies["NV_GiaoNhan"] != null ? Request.Cookies["NV_GiaoNhan"].Value : "";
            TDGH.Ly_do = Request.Cookies["Ly_Do"] != null ? Request.Cookies["Ly_Do"].Value : "";
            TDGH.Stt = Request.Cookies["STT"] != null ? Request.Cookies["STT"].Value : "";
            string result = "Error!";
            connectSQL();
            if (TDGH != null && TDGH.Details != null)
            {
                try
                {
                    var detailsTable = new DataTable();
                    detailsTable.Columns.Add("So_Hd", typeof(int));
                    detailsTable.Columns.Add("Ngay_HD", typeof(string));
                    detailsTable.Columns.Add("Ma_Dt", typeof(int));
                    detailsTable.Columns.Add("Ten_Dt", typeof(string));
                    detailsTable.Columns.Add("NV_GN", typeof(string));
                    detailsTable.Columns.Add("Giao_HD", typeof(bool));
                    detailsTable.Columns.Add("Tien", typeof(float));
                    detailsTable.Columns.Add("Tien_Phai_Thu", typeof(float));

                    detailsTable.Columns.Add("Noi_Dung", typeof(string));
                    detailsTable.Columns.Add("Chua_giao_hang", typeof(bool));
                    foreach (var detail in TDGH.Details)
                    {
                        detailsTable.Rows.Add(detail.So_Hd, detail.Ngay_HD, detail.Ma_Dt, detail.Ten_Dt, detail.NV_GiaoNhan, detail.Giao_HD, detail.Tien_HD,detail.Tien_Phai_Thu, detail.Noi_Dung, detail.Chua_giao_hang);
                    }

                    using (var connection = new SqlConnection(con.ConnectionString))
                    {
                        connection.Open();

                        using (var command = new SqlCommand("UpdateTheoDoiGiaoHang_SAP", connection))
                        {
                            command.CommandType = CommandType.StoredProcedure;

                            command.Parameters.AddWithValue("@_Ngay_Ct", TDGH.Ngay_Ct);
                            command.Parameters.AddWithValue("@_NV_GiaoHang", TDGH.Ma_NVGH);
                            command.Parameters.AddWithValue("@_Ten_NVGiaoHang", TDGH.Ten_NVGH);
                            command.Parameters.AddWithValue("@_Ten_NVPhuKho", TDGH.NV_GiaoNhan);
                            command.Parameters.AddWithValue("@_Ly_Do", TDGH.Ly_do);
                            command.Parameters.AddWithValue("@_Stt", TDGH.Stt);


                            // Pass details as a TVP parameter
                            var detailsParam = command.Parameters.AddWithValue("@_Details", detailsTable);
                            detailsParam.SqlDbType = SqlDbType.Structured;
                            detailsParam.TypeName = "B30GdetailType"; // Replace with your actual type name

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

        public ActionResult MauInGiaoHang_CNCT()
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
        private void FormatCellNoQH(ExcelRangeBase cell)
        {
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // Các định dạng khác của ô có thể thêm vào tại đây
        }
        private void FormatCell(ExcelRangeBase cell)
        {
            //cell.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // Các định dạng khác của ô có thể thêm vào tại đây
        }
        public List<TheoDoiGiaoHang> LoadDataTBNoQH()
        {
            connectSQL();
            List<TheoDoiGiaoHang> dataItems = new List<TheoDoiGiaoHang>();
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("MauInGiaoHang", connection))
                {
                    string Stt = Request.Cookies["stt"].Value;

                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;
                    using (SqlDataAdapter sda = new SqlDataAdapter(command))
                    {
                        DataSet ds = new DataSet();
                        command.Parameters.AddWithValue("@_Stt", Stt);
                        sda.Fill(ds);

                        // Kiểm tra xem DataSet có bảng dữ liệu hay không
                        if (ds.Tables.Count > 0)
                        {
                            DataTable dt = ds.Tables[2];

                            foreach (DataRow row in dt.Rows)
                            {
                                TheoDoiGiaoHang dataItem = new TheoDoiGiaoHang
                                {
                                    Ten_Dt = row["Ten_Dt"].ToString(),
                                    So_HD = row["So_Hd"].ToString(),
                                    Ngay_HD = row["Ngay_HD"].ToString(),
                                    Han_TT = row["Han_TT"].ToString(),
                                    Tien = row["Tien_Hd"].ToString(),
                                    Tien1 = row["Tien_Phai_Thu"].ToString(),






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
            var fileName = $"TheoDoiGiaoHang{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            // Lấy dữ liệu từ cookie
            List<TheoDoiGiaoHang> combinedData = LoadDataTBNoQH();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("MySheet");
                worksheet.View.ShowGridLines = false;

                // ... (Các bước tạo nội dung tệp Excel như bạn đã làm)
                // Đường dẫn đến hình ảnh trong thư mục 'image'
                var imagePath = Server.MapPath("~/assets/images/logo.png"); // Thay thế bằng đường dẫn thật
                                                                            // Lấy giá trị từ biến Dvcs
                string Year = Request.Cookies["namCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["namCookie"].Value) : "";
                string Month = Request.Cookies["thangCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["thangCookie"].Value) : "";
                string TenNV = Request.Cookies["tenNVCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tenNVCookie"].Value) : "";
                string SoCT = Request.Cookies["soCtCCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["soCtCCookie"].Value) : "";
                string Day = Request.Cookies["ngayCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["ngayCookie"].Value) : "";
                string TuyenGH = Request.Cookies["tenNVPhuCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tenNVPhuCookie"].Value) : "";
                string Sum = Request.Cookies["tongCongCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tongCongCookie"].Value) : "";
                string Sum1= Request.Cookies["tongCongCookie1"] != null ? HttpUtility.UrlDecode(Request.Cookies["tongCongCookie1"].Value) : "";

                // Đặt font chữ "Arial" cho toàn bộ tệp Excel
                worksheet.Cells.Style.Font.Name = "Times New Roman";

                // Chèn hình ảnh từ tệp hình vào ô A1
                ExcelPicture picture = worksheet.Drawings.AddPicture("MyPicture", new FileInfo(imagePath));
                picture.SetSize(75, 60); // Đặt kích thước cho hình ảnh
                picture.From.Row = 2;
                picture.From.Column = 1;
                worksheet.Column(1).Width = 10;



                worksheet.Cells["A1:I1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                worksheet.Cells["A1:A5"].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                worksheet.Cells["A5:I5"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                worksheet.Cells["J1:J5"].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                worksheet.Cells["B1:B5"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                worksheet.Cells["H1:H5"].Style.Border.Left.Style = ExcelBorderStyle.Medium;

                // Đặt văn bản vào ô A2
                worksheet.Cells["A1"].Value = "CTY CỔ PHẦN DƯỢC PHẨM";
                worksheet.Cells["B2"].Value = "OPC";
                worksheet.Cells["B2"].Style.Font.Bold = true;
                worksheet.Cells["A1"].Style.Font.Size = 13;
                worksheet.Cells["B2"].Style.Font.Size = 13;
                var cellB1 = worksheet.Cells["A1"];
                cellB1.Style.Font.Bold = true;

                worksheet.Cells["B2"].Style.Indent = 1;

                worksheet.Cells["H2"].Value = "PHỤ LỤC 2";
                worksheet.Cells["H2"].Style.Indent = 4;
                worksheet.Cells["H2"].Style.Font.Bold = true;
                worksheet.Cells["H2"].Style.Font.Size = 13;
                worksheet.Cells["H3"].Value = "BFO.638.1";
                worksheet.Cells["H3"].Style.Indent = 4;
                worksheet.Cells["H3"].Style.Font.Bold = true;
                worksheet.Cells["H3"].Style.Font.Size = 13;

                worksheet.Cells["C2"].Value = "PHIẾU ĐIỀU PHỐI GIAO HÀNG";
                //worksheet.Cells["D2"].Style.Indent = 3;
                worksheet.Cells["C2"].Style.Font.Bold = true;
                worksheet.Cells["C2"].Style.Font.Size = 16;
                worksheet.Cells["C2"].Style.Indent = 3;
                worksheet.Cells["C3"].Value = $"Số: {SoCT}, ngày {Day} tháng {Month} năm {Year}";
                worksheet.Cells["C3"].Style.Indent = 1;
                worksheet.Cells["C3"].Style.Font.Bold = true;
                worksheet.Cells["C3"].Style.Font.Size = 13;

                worksheet.Cells["D6"].Value = "CHI NHÁNH CẦN THƠ";
                worksheet.Cells["D6"].Style.Font.Bold = true;
                worksheet.Cells["D6"].Style.Font.Size = 15;

                worksheet.Cells["A6:I6"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                worksheet.Cells["I6"].Style.Border.Right.Style = ExcelBorderStyle.Medium;

                worksheet.Cells["A7"].Value = $"Người giao hàng - thu tiền: {TenNV}";
                worksheet.Cells["A7"].Style.Font.Bold = true;
                worksheet.Cells["A7"].Style.Font.Size = 15;

                worksheet.Cells["A8"].Value = $"Tuyến giao hàng: {TuyenGH}:(GN)";
                worksheet.Cells["A8"].Style.Font.Bold = true;
                worksheet.Cells["A8"].Style.Font.Size = 15;

                var startRow = 11;
                var startColumn = 1;

                var sttCell = worksheet.Cells[startRow - 1, startColumn, startRow, startColumn];
                sttCell.Value = "STT";
                sttCell.Merge = true;
                sttCell.Style.Font.Bold = true;
                sttCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                sttCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 240)); // Sử dụng hằng số màu sắc
                sttCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);


                var donViCell = worksheet.Cells[startRow - 1, startColumn + 1, startRow, startColumn + 1];
                donViCell.Value = "Đơn Vị";
                donViCell.Merge = true;
                donViCell.Style.Font.Bold = true;
                donViCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                donViCell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFFF0"));
                donViCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                var hoaDonCell = worksheet.Cells[startRow - 1, startColumn + 2, startRow - 1, startColumn + 5];
                hoaDonCell.Value = "Hóa Đơn";
                hoaDonCell.Merge = true;
                hoaDonCell.Style.Font.Bold = true;
                hoaDonCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                hoaDonCell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFFF0"));
                hoaDonCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                var soCell = worksheet.Cells[startRow, startColumn + 2];
                soCell.Value = "Số";
                soCell.Style.Font.Bold = true;
                soCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                soCell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFFF0"));
                soCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                soCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                soCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                var ngayCell = worksheet.Cells[startRow, startColumn + 3];
                ngayCell.Value = "Ngày";
                ngayCell.Style.Font.Bold = true;
                ngayCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                ngayCell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFFF0"));
                ngayCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ngayCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ngayCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                var tienHDCell = worksheet.Cells[startRow, startColumn + 4];
                tienHDCell.Value = "Tiền HĐ";
                tienHDCell.Style.Font.Bold = true;
                tienHDCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                tienHDCell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFFF0"));
                tienHDCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                tienHDCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                tienHDCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                var tienThuCell = worksheet.Cells[startRow, startColumn + 5];
                tienThuCell.Value = "Tiền Thu";
                tienThuCell.Style.Font.Bold = true;
                tienThuCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                tienThuCell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFFF0"));
                tienThuCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                tienThuCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                tienThuCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                var tongDVCell = worksheet.Cells[startRow - 1, startColumn + 6, startRow - 1, startColumn + 7];
                tongDVCell.Value = "Tổng Đơn Vị";
                tongDVCell.Merge = true;
                tongDVCell.Style.Font.Bold = true;
                tongDVCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                tongDVCell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFFF0"));
                tongDVCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                var thungCell = worksheet.Cells[startRow, startColumn + 6];
                thungCell.Value = "Thùng";
                thungCell.Style.Font.Bold = true;
                thungCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                thungCell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFFF0"));
                thungCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                thungCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                thungCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                var leCell = worksheet.Cells[startRow, startColumn + 7];
                leCell.Value = "Lẻ";
                leCell.Style.Font.Bold = true;
                leCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                leCell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFFF0"));
                leCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                leCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                leCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                var hanMucTTCell = worksheet.Cells[startRow - 1, startColumn + 8, startRow, startColumn + 8];
                hanMucTTCell.Value = "Hạn Mức TT";
                hanMucTTCell.Merge = true;
                hanMucTTCell.Style.Font.Bold = true;
                hanMucTTCell.Style.WrapText = true;
                hanMucTTCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                hanMucTTCell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFFF0"));
                hanMucTTCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                hanMucTTCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                hanMucTTCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                for (int col = 0; col < 8; col++)
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
                //var columnHeaderStyle = worksheet.Cells[startRow - 1, startColumn, startRow - 1, startColumn + 8].Style;
                //columnHeaderStyle.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black); // Đóng khung solid đen

                worksheet.Column(startColumn + 1).Width = 25; // Độ rộng cột cho "Tên sản phẩm"
                worksheet.Row(startRow - 1).Height = 25;
                worksheet.Row(startRow).Height = 25;

                worksheet.Column(startColumn + 2).Width = 10;
                worksheet.Column(startColumn + 3).Width = 12;
                worksheet.Column(startColumn + 4).Width = 10;
                worksheet.Column(startColumn + 5).Width = 12;
                worksheet.Column(startColumn + 6).Width = 12;
                worksheet.Column(startColumn + 7).Width = 12;
                worksheet.Column(startColumn + 8).Width = 10;
                worksheet.Column(startColumn + 11).Width = 20;
                worksheet.Column(startColumn + 12).Width = 15;


                //Đảm bảo rằng có dữ liệu trong biến tableData
                if (combinedData != null && combinedData.Any())
                {
                    var stt = 1;
                    var start = startRow + 1;
                    string previousTenDt = null; // Biến để lưu giá trị Ten_Dt của hàng trước đó

                    // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
                    for (int row = 0; row < combinedData.Count; row++)
                    {
                        var rowData = combinedData[row];

                        FormatCellNoQH(worksheet.Cells[start + row, startColumn]);

                        worksheet.Cells[start + row, startColumn].Value = stt;

                        // Kiểm tra giá trị của Ten_Dt so với hàng trước đó
                        if (rowData.Ten_Dt != previousTenDt)
                        {
                            worksheet.Cells[start + row, startColumn + 1].Value = rowData.Ten_Dt;
                            FormatCell(worksheet.Cells[start + row, startColumn + 1]);
                            previousTenDt = rowData.Ten_Dt; // Cập nhật giá trị mới
                        }
                        else
                        {
                            worksheet.Cells[start + row, startColumn + 1].Value = null; // Để trống nếu giống với hàng trước
                            //FormatCellNoQH(worksheet.Cells[start + row, startColumn + 1]);
                        }

                        worksheet.Cells[start + row, startColumn + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        worksheet.Cells[start + row, startColumn + 1].Style.WrapText = true;

                        // Các cột khác bạn xử lý như bình thường
                        FormatCellNoQH(worksheet.Cells[start + row, startColumn + 2]);
                        worksheet.Cells[start + row, startColumn + 2].Value = rowData.So_HD;

                        FormatCellNoQH(worksheet.Cells[start + row, startColumn + 3]);
                        worksheet.Cells[start + row, startColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[start + row, startColumn + 3].Value = rowData.Ngay_HD;

                        worksheet.Cells[start + row, startColumn + 4].Value = rowData.Tien;
                        worksheet.Cells[start + row, startColumn + 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        FormatCellNoQH(worksheet.Cells[start + row, startColumn + 4]);

                        worksheet.Cells[start + row, startColumn + 5].Value = rowData.Tien1;
                        worksheet.Cells[start + row, startColumn + 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        FormatCellNoQH(worksheet.Cells[start + row, startColumn + 5]);

                        FormatCellNoQH(worksheet.Cells[start + row, startColumn + 6]);
                        FormatCellNoQH(worksheet.Cells[start + row, startColumn + 7]);
                        FormatCellNoQH(worksheet.Cells[start + row, startColumn + 8]);

                        worksheet.Cells[start + row, startColumn + 8].Value = rowData.Han_TT;
                        worksheet.Cells[start + row, startColumn + 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        stt++;
                    }
                }
                else
                {
                    worksheet.Cells[startRow, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
                }

                // Xóa hàng tiêu đề mặc định

                var dataRowStyle = worksheet.Cells[startRow, startColumn, startRow, startColumn + 5].Style;
                dataRowStyle.Font.Bold = false;
                dataRowStyle.Font.Color.SetColor(Color.Black);
                dataRowStyle.Fill.PatternType = ExcelFillStyle.None;
                // Tạo bảng trong tệp Excel
                var rowTC = startRow + combinedData.Count;
                var TCCell = worksheet.Cells[rowTC + 1, startColumn, rowTC + 1, startColumn + 3];
                TCCell.Merge = true;
                TCCell.Value = "Tổng Cộng:";
                TCCell.Style.Font.Bold = true;
                TCCell.Style.Fill.PatternType = ExcelFillStyle.Solid; // Đặt kiểu nền
                TCCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFFF0"));
                FormatCellNoQH(TCCell);
                worksheet.Row(rowTC + 1).Height = 25;
                var sumCell = worksheet.Cells[rowTC + 1, startColumn + 4];
                sumCell.Value = Sum;
                sumCell.Style.Font.Bold = true;
                sumCell.Style.Fill.PatternType = ExcelFillStyle.Solid; // Đặt kiểu nền
                sumCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFFF0"));
                FormatCellNoQH(sumCell);
                

                var sum1Cell = worksheet.Cells[rowTC + 1, startColumn + 5];
                sum1Cell.Value = Sum1;
                sum1Cell.Style.Font.Bold = true;
                sum1Cell.Style.Fill.PatternType = ExcelFillStyle.Solid; // Đặt kiểu nền
                sum1Cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFFF0"));
                FormatCellNoQH(sum1Cell);
                
                FormatCellNoQH(worksheet.Cells[rowTC + 1, startColumn + 6]);
                worksheet.Cells[rowTC + 1, startColumn + 6].Style.Fill.PatternType = ExcelFillStyle.Solid; // Đặt kiểu nền
                worksheet.Cells[rowTC + 1, startColumn + 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFFF0"));
                FormatCellNoQH(worksheet.Cells[rowTC + 1, startColumn + 7]);
                worksheet.Cells[rowTC + 1, startColumn + 7].Style.Fill.PatternType = ExcelFillStyle.Solid; // Đặt kiểu nền
                worksheet.Cells[rowTC + 1, startColumn + 7].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFFF0"));
                FormatCellNoQH(worksheet.Cells[rowTC + 1, startColumn + 8]);
                worksheet.Cells[rowTC + 1, startColumn + 8].Style.Fill.PatternType = ExcelFillStyle.Solid; // Đặt kiểu nền
                worksheet.Cells[rowTC + 1, startColumn + 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFFF0"));
                string currentDate = DateTime.Now.ToString("HH:mm dd/MM/yyyy");
                var endRow = rowTC + 2;
                worksheet.Cells[endRow,startColumn ].Value = currentDate;
                worksheet.Cells[endRow, startColumn].Style.Font.Bold = true;
                worksheet.Cells[endRow, startColumn].Style.Font.Size = 12;
                worksheet.Row(endRow+1).Height = 30;
                worksheet.Cells[endRow+1, startColumn, endRow+1, startColumn + 8].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[endRow +1, startColumn].Value = "Đánh giá Nghiệp vụ kho: ";
                worksheet.Cells[endRow+1, startColumn].Style.Font.Bold = true;
                worksheet.Cells[endRow + 1, startColumn].Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Căn lên trên cùng
                worksheet.Cells[endRow + 2, startColumn].Value = "Báo cáo Tình hình giao hàng: ";
                worksheet.Cells[endRow + 2, startColumn].Style.Font.Bold = true;
                worksheet.Row(endRow + 2).Height = 75;
                worksheet.Cells[endRow + 2, startColumn].Style.VerticalAlignment = ExcelVerticalAlignment.Top; // Căn lên trên cùng
                worksheet.Cells[endRow + 2, startColumn, endRow + 2, startColumn + 8].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[endRow + 2, startColumn+8, endRow + 2, startColumn + 8].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                int nextRow = endRow + 3;
                worksheet.Cells[nextRow, startColumn].Value = "1.Chất lượng sản phẩm";
                worksheet.Cells[nextRow + 1, startColumn].Value = "- SP đang lưu hành, chất lượng tốt................, Hàng đổi trả................, Bể vỡ..................., Thu hồi...............";
                worksheet.Cells[nextRow + 2, startColumn].Value = "2.Điều kiện bảo quản";
                worksheet.Cells[nextRow + 3, startColumn].Value = "- Nhiệt độ thường ≤ 30 độ C............ ";
                worksheet.Cells[nextRow + 4, startColumn].Value = "-Điều kiện khác:....................................................................................................................";
                worksheet.Cells[nextRow + 4, startColumn, nextRow + 4, startColumn+8].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[nextRow, startColumn + 8, nextRow + 4, startColumn + 8].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[nextRow + 6, startColumn + 1].Value = "Điều Phối Giao Hàng";
                worksheet.Cells[nextRow + 6, startColumn + 1].Style.Font.Bold = true;


                worksheet.Cells[nextRow + 6, startColumn + 4].Value = "Thủ Kho";
                worksheet.Cells[nextRow + 6, startColumn + 4].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 6, startColumn + 8].Value = "Người Giao Hàng";
                worksheet.Cells[nextRow + 6, startColumn + 8].Style.Font.Bold = true;


                package.Save();
                byte[] fileBytes = package.GetAsByteArray();

                // Trả về tệp Excel dưới dạng dữ liệu binary
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

            }



            return View("MauInGiaoHang_CNCT");
        }

    }
}