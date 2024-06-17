using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using web4.Models;

namespace web4.Controllers
{
    public class Connect : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        //
        // GET: /test/
        public ActionResult Index()
        {
            return View();
        }
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "B7_OPC" + ";uid=sa;password=Hai@thong";
        }
    }
}