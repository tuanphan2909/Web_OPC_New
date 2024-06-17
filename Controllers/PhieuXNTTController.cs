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
    public class PhieuXNTTController : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: PhieuXNTT

        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
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
        private void FormatCellNoQH(ExcelRangeBase cell)
        {
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // Các định dạng khác của ô có thể thêm vào tại đây
        }
        public List<GetData> LoadDataPhieuNhapXNTT()
        {
            connectSQL();
            List<GetData> dataItems = new List<GetData>();
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_XacNhanCKTT_SAP]", connection))
                {
                    var SoCT = Request.Cookies["So_Ct"] != null ? Request.Cookies["So_Ct"].Value : "";
                    var fromDate = Request.Cookies["From_date"].Value;
                    var toDate = Request.Cookies["To_Date"].Value;
                    var MaDt = Request.Cookies["Ma_DT"] != null ? Request.Cookies["Ma_DT"].Value : string.Empty;
                    var ma_dvcs = Request.Cookies["MA_DVCS"].Value;
                    var ma_dt = Request.Cookies["Ma_Dt"].Value;
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;
                    using (SqlDataAdapter sda = new SqlDataAdapter(command))
                    {
                        DataSet ds = new DataSet();
                        command.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                        command.Parameters.AddWithValue("@_Den_Ngay", toDate);
                        command.Parameters.AddWithValue("@_Ma_dt", MaDt);
                        command.Parameters.AddWithValue("@_so_Ct", SoCT);
                        command.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                        sda.Fill(ds);

                        // Kiểm tra xem DataSet có bảng dữ liệu hay không
                        if (ds.Tables.Count > 0)
                        {
                            DataTable dt = ds.Tables[0];
                            foreach (DataRow row in dt.Rows)
                            {
                                GetData dataItem = new GetData
                                {
                                    So = row["So_Ct"].ToString(),
                                    NgayHD = row["Ngay_Ct1"].ToString(),
                                    NgayDenHan = row["Han_Thanh_Toan"].ToString(),
                                    TienTT = row["Tien_Truoc_Thue"].ToString(),
                                    TienHD = row["Tong_Tien"].ToString(),
                                    CKTT = row["CKTT"].ToString(),
                                    TienThanhToan = row["Tien_TT"].ToString(),





                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }
            return dataItems;
        }
        public ActionResult ExportPhieuNhapXNTT()
        {
            var fileName = $"PhieuNhapXNTT{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            // Lấy dữ liệu từ cookie
            List<GetData> combinedData = LoadDataPhieuNhapXNTT();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("MySheet");
                worksheet.View.ShowGridLines = false;

                // ... (Các bước tạo nội dung tệp Excel như bạn đã làm)
                // Đường dẫn đến hình ảnh trong thư mục 'image'


                string diachi = Request.Cookies["diachiCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["diachiCookie"].Value) : "";
                string ten_dt = Request.Cookies["tenDTCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tenDTCookie"].Value) : "";
                string ten_tdv = Request.Cookies["tenTDVCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tenTDVCookie"].Value) : "";

                string TienTT = Request.Cookies["tienTTCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tienTTCookie"].Value) : "";
                string tongTien = Request.Cookies["tongTienCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tongTienCookie"].Value) : "";
                string CKTT = Request.Cookies["CKTTCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["CKTTCookie"].Value) : "";
                string tongTienTT = Request.Cookies["tongTienTTCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tongTienTTCookie"].Value) : "";
                string ngay = Request.Cookies["ngay"] != null ? HttpUtility.UrlDecode(Request.Cookies["ngay"].Value) : "";
                string thang = Request.Cookies["thang"] != null ? HttpUtility.UrlDecode(Request.Cookies["thang"].Value) : "";
                string nam = Request.Cookies["nam"] != null ? HttpUtility.UrlDecode(Request.Cookies["nam"].Value) : "";
                string fromDate = Request.Cookies["fromIn"] != null ? HttpUtility.UrlDecode(Request.Cookies["fromIn"].Value) : "";
                string toDate = Request.Cookies["ToIn"] != null ? HttpUtility.UrlDecode(Request.Cookies["toIn"].Value) : "";
                // Đặt font chữ "Arial" cho toàn bộ tệp Excel
                worksheet.Cells.Style.Font.Name = "Times New Roman";


                // Đặt văn bản vào ô A2
                worksheet.Cells["A1"].Value = "CÔNG TY CỔ PHẦN DƯỢC PHẨM OPC";
                var cellB1 = worksheet.Cells["A1"];
                cellB1.Style.Font.Bold = true;

                worksheet.Cells["D3"].Value = "XÁC NHẬN THANH TOÁN";
                worksheet.Cells["D3"].Style.Font.Bold = true;
                worksheet.Cells["D3"].Style.Font.Size = 16;
                worksheet.Cells["D4"].Value = $"Từ {fromDate} đến {toDate}";
                worksheet.Cells["D4"].Style.Indent = 4;
                worksheet.Cells["A6"].Value = "- Căn cứ Luật thương mại số 36/2005/QH11 ngày 14/06/2005;";
                worksheet.Cells["A6"].Style.Font.Bold = true;
                worksheet.Cells["A7"].Value = "-Căn cứ vào hợp đồng số: ................../thỏa thuận của đôi bên.";
                worksheet.Cells["A7"].Style.Font.Bold = true;
                worksheet.Cells["A8"].Value = $"Khách hàng: {ten_dt} Tên trình dược viên: {ten_tdv}";
                worksheet.Cells["A9"].Value = $"Địa chỉ: {diachi}";
                worksheet.Cells["A10"].Value = $"Điều kiện thanh toán: ";
                worksheet.Cells["A11"].Value = $"Khi thanh toán ngay, Quý khách hàng được chiết khấu 2% trên giá trị thanh toán của những hóa đơn cụ thể như sau:";
                var startRow = 13;
                var startColumn = 1;
                worksheet.Cells[startRow - 1, startColumn].Value = "SỐ";
                worksheet.Cells[startRow - 1, startColumn + 1].Value = "NGÀY HÓA ĐƠN";
                worksheet.Cells[startRow - 1, startColumn + 2].Value = "NGÀY ĐẾN HẠN";
                worksheet.Cells[startRow - 1, startColumn + 3].Value = "TIỀN TRƯỚC THUẾ";
                worksheet.Cells[startRow - 1, startColumn + 4].Value = "TIỀN HĐ";
                worksheet.Cells[startRow - 1, startColumn + 5].Value = "CKTT";
                worksheet.Cells[startRow - 1, startColumn + 6].Value = "TIỀN THANH TOÁN";
                for (int col = 0; col < 7; col++)
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
                var columnHeaderStyle = worksheet.Cells[startRow - 1, startColumn, startRow - 1, startColumn + 6].Style;
                columnHeaderStyle.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black); // Đóng khung solid đen
                worksheet.Column(startColumn).Width = 18; // Độ rộng cột cho "STT"
                worksheet.Column(startColumn + 1).Width = 18; // Độ rộng cột cho "SỐ HÓA ĐƠN"
                worksheet.Column(startColumn + 2).Width = 18; // Độ rộng cột cho "NGÀY XUẤT"
                worksheet.Column(startColumn + 3).Width = 18; // Độ rộng cột cho "TIỀN NỢ"
                worksheet.Column(startColumn + 4).Width = 18; // Độ rộng cột cho "HẠN THANH TOÁN"
                worksheet.Column(startColumn + 5).Width = 18; // 
                worksheet.Column(startColumn + 6).Width = 18; // 

                // Đảm bảo rằng có dữ liệu trong biến tableData
                if (combinedData != null && combinedData.Any())
                {
                    var stt = 1;
                    // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
                    for (int row = 0; row < combinedData.Count; row++)
                    {
                        var rowData = combinedData[row];

                        worksheet.Cells[startRow + row, startColumn].Value = rowData.So;

                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn]);
                        worksheet.Cells[startRow + row, startColumn + 1].Value = rowData.NgayHD;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 1]);
                        worksheet.Cells[startRow + row, startColumn + 2].Value = rowData.NgayDenHan;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 2]);
                        worksheet.Cells[startRow + row, startColumn + 3].Value = rowData.TienTT;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 3]);
                        worksheet.Cells[startRow + row, startColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[startRow + row, startColumn + 4].Value = rowData.TienHD;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 4]);
                        worksheet.Cells[startRow + row, startColumn + 5].Value = rowData.CKTT;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 5]);
                        worksheet.Cells[startRow + row, startColumn + 6].Value = rowData.TienThanhToan;
                        FormatCellNoQH(worksheet.Cells[startRow + row, startColumn + 6]);
                        stt++;
                    }

                }
                else
                {
                    worksheet.Cells[startRow, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
                }
                worksheet.Cells[startRow + combinedData.Count, startColumn + 1].Value = "Tổng cộng";
                worksheet.Cells[startRow + combinedData.Count, startColumn + 1].Style.Font.Bold = true;
                worksheet.Cells[startRow + combinedData.Count, startColumn + 3].Value = $"{TienTT}"; // Ví dụ: Ghi giá trị tổng vào cột thứ 4
                worksheet.Cells[startRow + combinedData.Count, startColumn + 3].Style.Font.Bold = true;
                worksheet.Cells[startRow + combinedData.Count, startColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[startRow + combinedData.Count, startColumn + 4].Value = $"{tongTien}"; // Ví dụ: Ghi giá trị tổng vào cột thứ 4
                worksheet.Cells[startRow + combinedData.Count, startColumn + 4].Style.Font.Bold = true;
                worksheet.Cells[startRow + combinedData.Count, startColumn + 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[startRow + combinedData.Count, startColumn + 5].Value = $"{CKTT}"; // Ví dụ: Ghi giá trị tổng vào cột thứ 4
                worksheet.Cells[startRow + combinedData.Count, startColumn + 5].Style.Font.Bold = true;
                worksheet.Cells[startRow + combinedData.Count, startColumn + 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[startRow + combinedData.Count, startColumn + 6].Value = $"{tongTienTT}"; // Ví dụ: Ghi giá trị tổng vào cột thứ 4
                worksheet.Cells[startRow + combinedData.Count, startColumn + 6].Style.Font.Bold = true;
                worksheet.Cells[startRow + combinedData.Count, startColumn + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[startRow + combinedData.Count, startColumn].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count, startColumn + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count, startColumn + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count, startColumn + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count, startColumn + 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count, startColumn + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Cells[startRow + combinedData.Count, startColumn + 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                // Xóa hàng tiêu đề mặc định

                var dataRowStyle = worksheet.Cells[startRow, startColumn, startRow, startColumn + 6].Style;
                dataRowStyle.Font.Bold = false;
                dataRowStyle.Font.Color.SetColor(Color.Black);
                dataRowStyle.Fill.PatternType = ExcelFillStyle.None;
                // Tạo bảng trong tệp Excel
                var endRow = startRow + combinedData.Count;


                //var tableRange = worksheet.Cells[startRow, startColumn, endRow, endColumn];
                //var table = worksheet.Tables.Add(tableRange, "MyTable");
                //table.TableStyle = TableStyles.Light1;
                int nextRow = endRow + 2;
                worksheet.Cells[nextRow, startColumn].Value = $"Khách hàng đã thanh toán số tiền:........................, nhận chiết khấu thanh toán số tiền:......................";

                worksheet.Cells[nextRow + 2, startColumn + 6].Value = $"Ngày {ngay} tháng {thang} năm {nam}";
                worksheet.Cells[nextRow + 2, startColumn + 6].Style.Indent = 4;
                worksheet.Cells[nextRow + 2, startColumn + 6].Style.Font.Italic = true;
                worksheet.Cells[nextRow + 3, startColumn + 1].Value = "Khách Hàng";
                worksheet.Cells[nextRow + 3, startColumn + 1].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 3, startColumn + 4].Value = "Nhân viên thu tiền";
                worksheet.Cells[nextRow + 3, startColumn + 4].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 3, startColumn + 6].Value = "Người lập bảng";
                worksheet.Cells[nextRow + 3, startColumn + 6].Style.Font.Bold = true;
                worksheet.Cells[nextRow + 3, startColumn + 6].Style.Indent = 5;
                worksheet.Cells[nextRow + 4, startColumn].Value = "(Ký, ghi rõ họ tên)";
                worksheet.Cells[nextRow + 4, startColumn].Style.Indent = 8;
                worksheet.Cells[nextRow + 4, startColumn].Style.Font.Italic = true;

                package.Save();
                byte[] fileBytes = package.GetAsByteArray();

                // Trả về tệp Excel dưới dạng dữ liệu binary
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

            }




        }
    }
}