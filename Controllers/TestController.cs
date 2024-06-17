using System.Web.Mvc;
using System.Data.SqlClient;
using System.Data;
using System.Collections.Generic;
using System;
public class TestController : Controller
{
        public ActionResult Index()
        {
            List<string> imageSources = new List<string>();

            // Thiết lập chuỗi kết nối
            string connectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                // Tạo câu lệnh SQL
                string sqlQuery = "SELECT Hinh_Anh FROM B30ViengThamKH";

                // Tạo đối tượng SqlCommand
                SqlCommand cmd = new SqlCommand(sqlQuery, con);

                // Thực thi câu lệnh và nhận dữ liệu vào một đối tượng SqlDataReader
                SqlDataReader reader = cmd.ExecuteReader();

            // Duyệt qua từng dòng dữ liệu
            while (reader.Read())
            {
                string base64String = reader["Hinh_Anh"].ToString();
                string imgSrc = $"data:image/GIF;base64,{base64String}";
                imageSources.Add(imgSrc);
            }

            // Đóng đối tượng đọc dữ liệu
            reader.Close();
            }

            // Truyền danh sách các đường dẫn hình ảnh đến view
            return View(imageSources);
        }
}
