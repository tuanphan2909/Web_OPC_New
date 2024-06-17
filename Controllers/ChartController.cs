using ASPNET_MVC_ChartsDemo.Models;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Web.Mvc;
using System;
using System.Linq;
using System.Web;
using web4.Models;
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
using System.Globalization;
using System.Configuration;


namespace ASPNET_MVC_ChartsDemo.Controllers
{
    public class ChartController : Controller
    {
        // GET: Home
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }
        public ActionResult Chart()
        {
            List<Chart> dataPoint = new List<Chart>();

            dataPoint.Add(new Chart("Samsung", 25));
            dataPoint.Add(new Chart("Micromax", 13));
            dataPoint.Add(new Chart("Lenovo", 8));
            dataPoint.Add(new Chart("Intex", 7));
            dataPoint.Add(new Chart("Reliance", (int)6.8));
            dataPoint.Add(new Chart("Others", (int)40.2));


            ViewBag.DataPoints = JsonConvert.SerializeObject(dataPoint);

            return View();
        }
        public ActionResult Demo()
        {
            List<B20DmDvt> dvcsList = GetDvcsData();
            return View(dvcsList);
        }
        public ActionResult DoanhThuChiNhanh_Admin()
        {
            List<Chart> dataPoint = new List<Chart>();
                
            dataPoint.Add(new Chart("Samsung", 25));
            dataPoint.Add(new Chart("Micromax", 13));
            dataPoint.Add(new Chart("Lenovo", 8));
            dataPoint.Add(new Chart("Intex", 7));
            dataPoint.Add(new Chart("Reliance", (int)6.8));
            dataPoint.Add(new Chart("Others",(int) 40.2));

            ViewBag.DataPoints = JsonConvert.SerializeObject(dataPoint);

            return View("DoanhThuChiNhanh_Admin");
        }
      
        public List<B20DmDvt> GetDvcsData()
        {
            List<B20DmDvt> dvcsList = new List<B20DmDvt>();

            try
            {
                using (SqlConnection connection = new SqlConnection("Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong"))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand("select * from B30ViengThamKH", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                B20DmDvt dvcs = new B20DmDvt
                                {
                                    tendangnhap = reader.GetString(1),       // Thay đổi index nếu cần
                                    hoten = reader.GetString(5),   // Thay đổi index nếu cần
                                   Hinh_Anh2 = reader.GetString(14) // Thay đổi index nếu cần
                                };

                                dvcsList.Add(dvcs);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Xử lý lỗi nếu cần
            }

            return dvcsList;
        }
    }
    //public ActionResult Index()
    //{
    //    List<B20DmDvt> images = GetImages();
    //    return View(images);
    //}

    //[HttpPost]
    //public ActionResult Index(int imageId)
    //{
    //    List<B20DmDvt> images = GetImages();
    //    GetImages image = images.Find(p => p.tendangnhap == imageId);
    //    if (image != null)
    //    {
    //        image.IsSelected = true;
    //        ViewBag.Base64String = "data:image/png;base64," + Convert.ToBase64String(image.Data, 0, image.Data.Length);
    //    }
    //    return View(images);
    //}

    //public List<B20DmDvt> GetImages()
    //{
    //    string query = "select * from B30ViengThamKH";
    //    List<B20DmDvt> images = new List<B20DmDvt>();
    //    string constr = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
    //    using (SqlConnection con = new SqlConnection(constr))
    //    {
    //        using (SqlCommand cmd = new SqlCommand(query))
    //        {
    //            cmd.CommandType = CommandType.Text;
    //            cmd.Connection = con;
    //            con.Open();
    //            using (SqlDataReader sdr = cmd.ExecuteReader())
    //            {
    //                while (sdr.Read())
    //                {
    //                    images.Add(new B20DmDvt
    //                    {
                           
    //                        Ten_Dvcs = sdr["Ma_Dvcs"].ToString(),
    //                        Hinh_Anh2 = sdr["Ten_Dt"].ToString(),
    //                        Hinh_Anh = (byte[])sdr["Hinh_Anh"]
    //                    });
    //                }
    //            }
    //            con.Close();
    //        }

    //        return images;
    //    }
    //    return View();
    //}

}