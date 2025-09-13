using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using MVCPRACTICES.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Configuration.Internal;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;

namespace MVCPRACTICES.Controllers
{
    public class HomeController : Controller
    {
        string connectionstring = ConfigurationManager.ConnectionStrings["my_connection"].ConnectionString;
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        //get method to insert student data
        public ActionResult StudentView()
        {
            return View();
        }
        //post method to insert student data
        public ActionResult InsertStudentData(STUDENTDATA_MODEL sdm)
        {
            try
            {
                using(SqlConnection sqlcon=new SqlConnection(connectionstring))
                {
                    SqlCommand cmd = new SqlCommand("USP_INSERTSTUDENTDATA", sqlcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@STUDENT_NAME", sdm.studentname);
                    cmd.Parameters.AddWithValue("@STUDENT_AGE", sdm.studentage);
                    cmd.Parameters.AddWithValue("@STUDENT_QUALIFICATION",sdm.studentqualification);
                    cmd.Parameters.AddWithValue("@STUDENT_GENDER",sdm.studentgender);
                    cmd.Parameters.AddWithValue("@COUNTRY_ID", sdm.countryid);
                    cmd.Parameters.AddWithValue("@STATE_ID",sdm.stateid);
                    sqlcon.Open();
                    int i = cmd.ExecuteNonQuery();
                    if (i > 0)
                    {
                        return Json(new { success = true, message = "Data Saved Successfully" },JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        return Json(new { success = false, message = "Unable to saved data" }, JsonRequestBehavior.AllowGet);
                    }
                }
            }
            catch(Exception ex)
            {
                throw ex;
                //return Json(new { success = false, message = "Internal server error" },JsonRequestBehavior.AllowGet);
            }
        }
        //show student data using online mode
        public ActionResult ShowStudentData(string StudentName,string Course)
        {
                int sr_no = 0;
                List<STUDENTDATA_MODEL> list = new List<STUDENTDATA_MODEL>();
                using(SqlConnection sqlcon=new SqlConnection(connectionstring))
                {
                    SqlCommand cmd = new SqlCommand("USP_READSTUDENDATA", sqlcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@STUDENT_NAME", StudentName);
                    cmd.Parameters.AddWithValue("@STUDENT_QUALIFICATION", Course);
                    sqlcon.Open();
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        sr_no = sr_no + 1;
                        STUDENTDATA_MODEL sdm = new STUDENTDATA_MODEL();
                        sdm.sr_no = sr_no;
                        sdm.id = Convert.ToInt32(rdr["ID"]);
                        sdm.studentname = rdr["STUDENT_NAME"].ToString();
                        sdm.studentage = Convert.ToInt32(rdr["STUDENT_AGE"]);
                        sdm.studentqualification = rdr["STUDENT_QUALIFICATION"].ToString();
                        sdm.studentgender = rdr["STUDENT_GENDER"].ToString();
                        sdm.countryname = rdr["COUNTRY_NAME"] == DBNull.Value ? null : rdr["COUNTRY_NAME"].ToString();
                        sdm.statename = rdr["STATENAME"] == DBNull.Value ? null : Convert.ToString(rdr["STATENAME"]);

                        sdm.countryid = rdr["COUNTRY_ID"] == DBNull.Value ? 0 : Convert.ToInt32(rdr["COUNTRY_ID"]);
                        sdm.stateid = rdr["STATE_ID"] == DBNull.Value ? 0 : Convert.ToInt32(rdr["STATE_ID"]);

                    list.Add(sdm);
                    }
                }
            //return View(list);
        return Json(new { success = true, Data = list }, JsonRequestBehavior.AllowGet);
        }

        //show data using datatable offline mode
        public ActionResult ShowStudentDataUsingDT(string StudentName, string Course)
        {
            DataTable dt = new DataTable();
            List<STUDENTDATA_MODEL> list = new List<STUDENTDATA_MODEL>();
            using (SqlConnection sqlcon = new SqlConnection(connectionstring))
            {
                SqlCommand cmd = new SqlCommand("USP_READSTUDENDATA", sqlcon);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@STUDENT_NAME", StudentName);
                cmd.Parameters.AddWithValue("@STUDENT_QUALIFICATION", Course);

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);   // Fill DataTable with result
            }

            if (dt.Rows.Count > 0)
            {
                //Serializable Object → Object ko JSON / XML / Binary me convert karna.
                //Deserializable Object → JSON / XML / Binary ko wapas Object me convert karna.

                string jsonData = JsonConvert.SerializeObject(dt);
                list = JsonConvert.DeserializeObject<List<STUDENTDATA_MODEL>>(jsonData);
                return Json(new { success = true, Data = list }, JsonRequestBehavior.AllowGet);
            }
            else {
                return Json(new { success = false, message ="no item found", Data = ""}, JsonRequestBehavior.AllowGet);
            }

        }//isce data table me lane ke liyea datbase ke column name aur model ka column name same hone chahiyea


        //Delete student data
        public ActionResult DeleteStudentData(int Id)
        {
            try
            {
                using (SqlConnection sqlcon = new SqlConnection(connectionstring))
                {
                    SqlCommand cmd = new SqlCommand("USP_DELETESTUDENDATA", sqlcon);
                    cmd.CommandType= CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@ID", Id);
                    sqlcon.Open();
                    int i = cmd.ExecuteNonQuery();
                    if (i > 0)
                    {
                        return Json(new { success = true, message = "Data deleted successfully" }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        return Json(new { success = false, message = "Unable to delete data" }, JsonRequestBehavior.AllowGet);
                    }
                }
            }
            catch(Exception ex)
            {
                throw ex;
             //   return Json(new { success = false, message = "Internal server error" }, JsonRequestBehavior.AllowGet);
            }
        }
        //update student data
        public ActionResult UpdateStudentData(string Id,string StudentName,string StudentAge,string StudentQual,string StudentGender,
            string CountryId,string StateId)
        {
            try
            {
                using (SqlConnection sqlcon = new SqlConnection(connectionstring))
                {
                    SqlCommand cmd = new SqlCommand("USP_UPDATESTUDENTDATA", sqlcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    //Handle null value
                    cmd.Parameters.AddWithValue("@ID", (object)Id ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@STUDENT_NAME", (object)StudentName ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@STUDENT_AGE", (object)StudentAge ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@STUDENT_QUALIFICATION", (object)StudentQual ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@STUDENT_GENDER", (object)StudentGender ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@COUNTRY_ID", (object)CountryId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@STATEID", (object)StateId ?? DBNull.Value);

                    //cmd.Parameters.AddWithValue("@ID", Id);
                    //cmd.Parameters.AddWithValue("@STUDENT_NAME", StudentName);
                    //cmd.Parameters.AddWithValue("@STUDENT_AGE", StudentAge);
                    //cmd.Parameters.AddWithValue("@STUDENT_QUALIFICATION", StudentQual);
                    //cmd.Parameters.AddWithValue("@STUDENT_GENDER", StudentGender);
                    //cmd.Parameters.AddWithValue("@COUNTRY_ID", CountryId);
                    //cmd.Parameters.AddWithValue("@STATEID", StateId);
                    sqlcon.Open();
                    int i = cmd.ExecuteNonQuery();
                    if (i > 0)
                    {
                      return  Json(new { success = true, message = "Data updated successfully" }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                      return  Json(new { success = false, message = "Unable to update data" }, JsonRequestBehavior.AllowGet);
                    }
                }
            }
            catch(Exception ex)
            {
                throw ex;
                //Json(new { success = false, message = "Server error" }, JsonRequestBehavior.AllowGet);
            }
        }
        //country master get method
        public ActionResult GetCountry()
        {
            List<countrymaster_model> list=new List<countrymaster_model>();
            using (SqlConnection sqlcon=new SqlConnection(connectionstring))
            {
                SqlCommand cmd = new SqlCommand("USP_COUNTRY_MASTER", sqlcon);
                cmd.CommandType= CommandType.StoredProcedure;
                sqlcon.Open();
                SqlDataReader rdr=cmd.ExecuteReader();
                while (rdr.Read())
                {
                    countrymaster_model cmm = new countrymaster_model();
                    cmm.countryid = Convert.ToInt32(rdr["COUNTRY_ID"]);
                    cmm.countryname = rdr["COUNTRY_NAME"].ToString();
                    list.Add(cmm);
                }
            }
            return Json(new { success = true, message = "data fetched successfully", data = list }, JsonRequestBehavior.AllowGet);
        }
        //state master get method
        public ActionResult GetState(int CountryId)
        {
            List<statemaster_model> list = new List<statemaster_model>();
            using (SqlConnection sqlcon = new SqlConnection(connectionstring))
            {
                SqlCommand cmd = new SqlCommand("USP_STATE_MASTER", sqlcon);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@COUNTRY_ID",CountryId);
                sqlcon.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    statemaster_model smm = new statemaster_model();
                    smm.stateid = Convert.ToInt32(rdr["STATE_ID"]);
                    smm.statename = rdr["STATE_NAME"].ToString();
                    smm.countryid= Convert.ToInt32(rdr["COUNTRY_ID"]);
                    list.Add(smm);
                }
            }
            return Json(new { success = true, message = "data fetched successfully", data = list }, JsonRequestBehavior.AllowGet);
        }

    //export to excel datatables data
    public ActionResult ExportToExcel(string StudentName, string Course)
    {
        List<STUDENTDATA_MODEL> students = new List<STUDENTDATA_MODEL>();
        using (SqlConnection sqlcon = new SqlConnection(connectionstring))
        {
            SqlCommand cmd = new SqlCommand("USP_READSTUDENDATA", sqlcon);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@STUDENT_NAME", string.IsNullOrEmpty(StudentName)? string.Empty : StudentName);
            cmd.Parameters.AddWithValue("@STUDENT_QUALIFICATION", string.IsNullOrEmpty(Course)? string.Empty : Course);
            sqlcon.Open();
            SqlDataReader rdr = cmd.ExecuteReader();
            int sr_no = 0;

            while (rdr.Read())
            {
                sr_no++;
                STUDENTDATA_MODEL sdm = new STUDENTDATA_MODEL
                {
                    sr_no = sr_no,
                    id = Convert.ToInt32(rdr["ID"]),
                    studentname = rdr["STUDENT_NAME"].ToString(),
                    studentage = Convert.ToInt32(rdr["STUDENT_AGE"]),
                    studentqualification = rdr["STUDENT_QUALIFICATION"].ToString(),
                    studentgender = rdr["STUDENT_GENDER"].ToString(),
                    countryname = rdr["COUNTRY_NAME"] == DBNull.Value ? null : rdr["COUNTRY_NAME"].ToString(),
                    statename = rdr["STATENAME"] == DBNull.Value ? null : Convert.ToString(rdr["STATENAME"]),
                    countryid = rdr["COUNTRY_ID"] == DBNull.Value ? 0 : Convert.ToInt32(rdr["COUNTRY_ID"]),
                    stateid = rdr["STATE_ID"] == DBNull.Value ? 0 : Convert.ToInt32(rdr["STATE_ID"])
                };
                students.Add(sdm);
            }
        }

        // Convert list to DataTable (ClosedXML ke liye easy)
        DataTable dt = new DataTable("Students");
        dt.Columns.AddRange(new DataColumn[7] {
        new DataColumn("Sr. No"),
        new DataColumn("Name"),
        new DataColumn("Age"),
        new DataColumn("Qualification"),
        new DataColumn("Gender"),
        new DataColumn("Country"),
        new DataColumn("State")
    });

        foreach (var s in students)
        {
            dt.Rows.Add(s.sr_no, s.studentname, s.studentage,
                        s.studentqualification, s.studentgender,
                        s.countryname, s.statename);
        }

        using (XLWorkbook wb = new XLWorkbook())
        {
            wb.Worksheets.Add(dt);
            using (MemoryStream stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                return File(stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "StudentData.xlsx");
            }
        }
    }

       //structure for bulk upload
       public ActionResult ExportStudentTemplate()
       {
            // Step 1: Create DataTable structure
            DataTable dt = new DataTable("StudentsTemplate");
            dt.Columns.Add("Sr. No", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Qualification", typeof(string));
            dt.Columns.Add("Gender", typeof(string));
            dt.Columns.Add("Country", typeof(string));
            dt.Columns.Add("State", typeof(string));
            

            // Step 3: Generate Excel using ClosedXML
           using (XLWorkbook wb = new XLWorkbook())
           {
                var ws = wb.Worksheets.Add(dt);
                ws.Table(0).ShowAutoFilter = true; // Filter enable karne ke liye

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "StudentTemplate.xlsx");
                }
           }
       }
        //Next Action Method for bulk uploading


    }
}


