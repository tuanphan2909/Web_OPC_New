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
namespace web4.Controllers
{
    public class BienBanTuKiemKeController : Controller
    {
        // GET: BienBanTuKiemKe
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
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
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }
        public ActionResult BienBanTuKiemKe_In()
        {
            List<BKHoaDonGiaoHang> dmDlistVT = LoadDmVt();
            ViewBag.DataVT = dmDlistVT;
            var fromDate = Request.Cookies["From_date"].Value;
            var toDate = Request.Cookies["To_Date"].Value;
            DataSet ds = new DataSet();
            connectSQL();
            var Ma_Vt = Request.Cookies["Ma_Vt"].Value;
            var Ma_Kho = Request.Cookies["Ma_Kho"].Value;
            var Ma_DV = Request.Cookies["Ma_DV"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_BienBanTuKiemKe_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;
                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                con.Open();
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", Ma_DV);
                    cmd.Parameters.AddWithValue("@_Ma_Kho", Ma_Kho);
                    cmd.Parameters.AddWithValue("@_Ma_Vt", Ma_Vt);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult BienBanTuKiemKe_Fill()
        {
            List<BKHoaDonGiaoHang> dmDlistVT = LoadDmVt();
            ViewBag.DataVT = dmDlistVT;
            return View();
        }
        private void FormatCellNoQH(ExcelRangeBase cell)
        {
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // Các định dạng khác của ô có thể thêm vào tại đây
        }
        public List<BienBanTuKiemKe> LoadDataTBNoQH()
        {
            connectSQL();
            List<BienBanTuKiemKe> dataItems = new List<BienBanTuKiemKe>();
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_BienBanTuKiemKe_SAP]", connection))
                {
                    var fromDate = Request.Cookies["From_date"].Value;
                    var toDate = Request.Cookies["To_Date"].Value;
                    var Ma_Vt = Request.Cookies["Ma_Vt"].Value;
                    var Ma_Kho = Request.Cookies["Ma_Kho"].Value;
                    var Ma_DV = Request.Cookies["Ma_DV"].Value;
                    var ma_dvcs = Request.Cookies["MA_DVCS"].Value;
                  
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;
                    using (SqlDataAdapter sda = new SqlDataAdapter(command))
                    {
                        DataSet ds = new DataSet();
                        command.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                        command.Parameters.AddWithValue("@_Den_Ngay", toDate);
                        command.Parameters.AddWithValue("@_ma_dvcs", Ma_DV);
                        command.Parameters.AddWithValue("@_Ma_Kho", Ma_Kho);
                        command.Parameters.AddWithValue("@_Ma_Vt", Ma_Vt);
                        sda.Fill(ds);

                        // Kiểm tra xem DataSet có bảng dữ liệu hay không
                        if (ds.Tables.Count > 0)
                        {
                            DataTable dt = ds.Tables[0];

                            foreach (DataRow row in dt.Rows)
                            {
                                BienBanTuKiemKe dataItem = new BienBanTuKiemKe
                                {
                                    Ma_SP = row["Ma_Vt"].ToString(),
                                    TenSP = row["Ten_Vt"].ToString(),
                                    So_Lo = row["So_Lo"].ToString(),

                                    Han_Dung = row["Han_Dung"].ToString(),
                                    Dvt = row["Dvt"].ToString(),
                                    Ton_CK = row["So_Luong_CK"].ToString(),
                                    QuyDoiThung = row["So_Luong_Quy_Doi"].ToString(),
                                    Hop_Le = row["So_Luong_quy_Doi_Hop"].ToString(),
                                    //SL_ThucTe = Convert.ToDecimal(row["Tong_No"].ToString()),
                                    //Thua = Convert.ToDecimal(row["Tong_No"].ToString()),
                                    //Thieu = Convert.ToDecimal(row["Tong_No"].ToString()),
                                    //NhanXet = Convert.ToDecimal(row["Tong_No"].ToString()),
                                    //NguyenNhan = Convert.ToDecimal(row["Tong_No"].ToString()),





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
            var fileName = $"BienBanTuKiemKe{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            // Lấy dữ liệu từ cookie
            List<BienBanTuKiemKe> combinedData = LoadDataTBNoQH();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("MySheet");
                worksheet.View.ShowGridLines = false;

                // ... (Các bước tạo nội dung tệp Excel như bạn đã làm)
                // Đường dẫn đến hình ảnh trong thư mục 'image'
                var imagePath = Server.MapPath("~/assets/images/logo.png"); // Thay thế bằng đường dẫn thật
                                                                            // Lấy giá trị từ biến Dvcs
                string Year = Request.Cookies["Year"] != null ? HttpUtility.UrlDecode(Request.Cookies["Year"].Value) : "";
                string Month = Request.Cookies["Month"] != null ? HttpUtility.UrlDecode(Request.Cookies["Month"].Value) : "";
                string MaKho = Request.Cookies["Ma_Kho"] != null ? HttpUtility.UrlDecode(Request.Cookies["Ma_Kho"].Value) : "";
              
                // Đặt font chữ "Arial" cho toàn bộ tệp Excel
                worksheet.Cells.Style.Font.Name = "Times New Roman";

                // Chèn hình ảnh từ tệp hình vào ô A1
                ExcelPicture picture = worksheet.Drawings.AddPicture("MyPicture", new FileInfo(imagePath));
                picture.SetSize(75, 60); // Đặt kích thước cho hình ảnh
                picture.From.Row = 2;
                picture.From.Column =1;
                worksheet.Column(1).Width = 10;



                worksheet.Cells["A1:M1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                worksheet.Cells["A1:A5"].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                worksheet.Cells["A5:M5"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                worksheet.Cells["M1:M5"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                worksheet.Cells["B1:B5"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                worksheet.Cells["K1:K5"].Style.Border.Left.Style = ExcelBorderStyle.Medium;

                // Đặt văn bản vào ô A2
                worksheet.Cells["A1"].Value = "CTY CỔ PHẦN DƯỢC PHẨM";
                worksheet.Cells["B2"].Value = "OPC";
                worksheet.Cells["B2"].Style.Font.Bold = true;
                worksheet.Cells["A1"].Style.Font.Size = 13;
                worksheet.Cells["B2"].Style.Font.Size = 13;
                var cellB1 = worksheet.Cells["A1"];
                cellB1.Style.Font.Bold = true;

                worksheet.Cells["B2"].Style.Indent =1;
           
                worksheet.Cells["L2"].Value = "PHỤ LỤC 2";
                worksheet.Cells["L2"].Style.Indent = 5;
                worksheet.Cells["L2"].Style.Font.Bold = true;
                worksheet.Cells["L2"].Style.Font.Size = 13;
                worksheet.Cells["L3"].Value = "DD0.617.6";
                worksheet.Cells["L3"].Style.Indent = 5;
                worksheet.Cells["L3"].Style.Font.Bold = true;
                worksheet.Cells["L3"].Style.Font.Size = 13;

                worksheet.Cells["D2"].Value = "BIÊN BẢN TỰ KIỂM KÊ THÁNG";
                worksheet.Cells["D2"].Style.Indent = 3;
                worksheet.Cells["D2"].Style.Font.Bold = true;
                worksheet.Cells["D2"].Style.Font.Size = 20;
                worksheet.Cells["F3"].Value = $"{Month}/{Year}";
                worksheet.Cells["F3"].Style.Indent = 5;
                worksheet.Cells["F3"].Style.Font.Bold = true;
                worksheet.Cells["F3"].Style.Font.Size = 18;


                worksheet.Cells["B7"].Value = $"Mã Kho: {MaKho}";
                worksheet.Cells["B7"].Style.Font.Bold = true;
                worksheet.Cells["B7"].Style.Font.Size = 15;
               
                var startRow = 10;
                var startColumn = 1;
                worksheet.Cells[startRow - 1, startColumn].Value = "Mã SP";
                worksheet.Cells[startRow - 1, startColumn + 1].Value = "Tên Sản Phẩm - Quy Cách";
                worksheet.Cells[startRow - 1, startColumn + 2].Value = "Số Lô";
                worksheet.Cells[startRow - 1, startColumn + 3].Value = "Hạn Dùng";
                worksheet.Cells[startRow - 1, startColumn + 4].Value = "ĐVT";
                worksheet.Cells[startRow - 1, startColumn + 5].Value = "Tồn Cuối Kỳ";
                worksheet.Cells[startRow - 1, startColumn + 6].Value = "Quy Đổi Thùng";
                worksheet.Cells[startRow - 1, startColumn + 7].Value = "Hộp Lẻ";
                worksheet.Cells[startRow - 1, startColumn + 8].Value = "SL Thực Tế";
                worksheet.Cells[startRow - 1, startColumn + 9].Value = "Thừa";
                worksheet.Cells[startRow - 1, startColumn + 10].Value = "Thiếu";
                worksheet.Cells[startRow - 1, startColumn + 11].Value = "Nhận Xét Chất Lượng";
                worksheet.Cells[startRow - 1, startColumn + 12].Value = "Nguyên Nhân";
                for (int col = 0; col < 14; col++)
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
                var columnHeaderStyle = worksheet.Cells[startRow - 1, startColumn, startRow - 1, startColumn + 12].Style;
                columnHeaderStyle.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black); // Đóng khung solid đen
         
                worksheet.Column(startColumn + 1).Width = 25; // Độ rộng cột cho "Tên sản phẩm"
                worksheet.Column(startColumn + 2).Width = 10;
                worksheet.Column(startColumn + 3).Width = 12;
                worksheet.Column(startColumn + 4).Width = 10;
                worksheet.Column(startColumn + 5).Width = 12;
                worksheet.Column(startColumn + 6).Width = 12;
                worksheet.Column(startColumn + 11).Width = 20;  
                worksheet.Column(startColumn + 12).Width = 15;
              

                // Đảm bảo rằng có dữ liệu trong biến tableData
                if (combinedData != null && combinedData.Any())
                {
                  
                    // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
                    for (int row = 0; row < combinedData.Count; row++)
                    {
                        var rowData = combinedData[row];

                     

                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn]);

                        worksheet.Cells[startRow + row, startColumn ].Value = rowData.Ma_SP;

                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 1]);
                        worksheet.Cells[startRow + row, startColumn + 1].Value = rowData.TenSP;
                        worksheet.Cells[startRow + row, startColumn + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 2]);
                        worksheet.Cells[startRow + row, startColumn + 2].Value = rowData.So_Lo;

                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 3]);
                        worksheet.Cells[startRow + row, startColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[startRow + row, startColumn + 3].Value = rowData.Han_Dung;

                        worksheet.Cells[startRow + row, startColumn + 4].Value = rowData.Dvt;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 4]);
                        worksheet.Cells[startRow + row, startColumn + 5].Value = rowData.Ton_CK;
                        worksheet.Cells[startRow + row, startColumn + 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 5]);
                        worksheet.Cells[startRow + row, startColumn + 6].Value = rowData.QuyDoiThung;
                        worksheet.Cells[startRow + row, startColumn + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 6]);
                        worksheet.Cells[startRow + row, startColumn + 7].Value = rowData.Hop_Le;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 7]);
                        worksheet.Cells[startRow + row, startColumn + 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 8]);
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 9]);
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 10]);
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 11]);
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 12]);

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
                var endRow = startRow + combinedData.Count;


                //var tableRange = worksheet.Cells[startRow, startColumn, endRow, endColumn];
                //var table = worksheet.Tables.Add(tableRange, "MyTable");
                //table.TableStyle = TableStyles.Light1;
                int nextRow = endRow + 1;
                worksheet.Cells[nextRow +1, startColumn + 10].Value = "Ngày........../.........../.............";
                worksheet.Cells[nextRow +2, startColumn].Style.Indent = 2;
                worksheet.Cells[nextRow +2, startColumn].Style.Font.Italic = true;
                worksheet.Cells[nextRow +2, startColumn + 2].Value = "Thủ Kho";
                worksheet.Cells[nextRow+2, startColumn + 2].Style.Font.Bold = true;
              
                worksheet.Cells[nextRow +2, startColumn + 10].Value = "Phụ Trách Đơn Vị Quản Kho";
                worksheet.Cells[nextRow +2, startColumn + 10].Style.Font.Bold = true;
                worksheet.Cells[nextRow +3, startColumn+2].Value = "(Ký, đóng dấu, ghi rõ họ tên)";
              
                worksheet.Cells[nextRow +3, startColumn+2].Style.Font.Italic = true;

                package.Save();
                byte[] fileBytes = package.GetAsByteArray();

                // Trả về tệp Excel dưới dạng dữ liệu binary
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

            }



            return View("BienBanTuKiemKe_In");
        }

    }
}