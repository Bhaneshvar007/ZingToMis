
using Dapper;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Linq.Mapping;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace ZingToMIS
{ 
    public class MISdbprocess
    {
        public static string strLogs = ConfigurationManager.AppSettings["Logs"];
        public static string constr = ConfigurationManager.AppSettings["ConnectionString"];
        public static string FilePath = ConfigurationManager.AppSettings["MailFilePath"];
        public static string MAILTO = ConfigurationManager.AppSettings["MAILTO"];
        public static string strEmailIPAddress = ConfigurationManager.AppSettings["EMAILIPADDRESS"];
        public void ActivityStart()
        {

            try
            {
                List<EMPCODE> EmpCodelist = new List<EMPCODE>();


                var result = "";
                var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://hrapi.utiamc.com/api/DataServices/GetEmployeeDetails");
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "POST";
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                Root root = new Root();
                List<MISDBModel> mISDBModel = new List<MISDBModel>();
                string FROMDATE = "";
                try
                {
                    using (OracleConnection conn = new OracleConnection(constr))
                    {
                        //conn.Open();
                        string query = "SELECT * FROM MISTEST.EMPLOYEE_LIST_FROM_ZING WHERE LST_UPD_DT = '"+ DateTime.Now.Date.AddDays(-2).ToString("dd-MMM-yyyy").ToUpper() +"' ";
                        //EmpCodelist = conn.Query<EMPCODE>(query).ToList();
                        if (EmpCodelist.Count == 0)
                        {
                            FROMDATE = DateTime.Now.Date.AddDays(-2).ToString("dd-MMM-yyyy").ToUpper();
                        }
                        else 
                        {
                            FROMDATE = DateTime.Now.Date.AddDays(-1).ToString("dd-MMM-yyyy").ToUpper();
                        }
                    }
                }
                catch (Exception ex)
                {
                    FROMDATE = DateTime.Now.Date.AddDays(-1).ToString("dd-MMM-yyyy").ToUpper();
                }
                // string TODATE = DateTime.Now.ToString("dd-MMM-yyyy").ToUpper();
                string TODATE = DateTime.Now.Date.AddDays(-1).ToString("dd-MMM-yyyy").ToUpper();
                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    //string json = "{\"ApplicationCode\":\"\"," +
                    //               "\"Token\":\"65EBEF06511C4E6CB09A94AFA99DA439\" } ";

                    //string json = new JavaScriptSerializer().Serialize(new
                    //{
                    //    ApplicationCode = "MISVPAY",
                    //    Token = "65EBEF06511C4E6CB09A94AFA99DA439",
                    //    From_Date = "14-DEC-2023",
                    //    To_Date = "31-DEC-2023"
                    //});

                    string json = new JavaScriptSerializer().Serialize(new
                    {
                        ApplicationCode = "MISVPAY",
                        Token = "65EBEF06511C4E6CB09A94AFA99DA439"
                        //From_Date = FROMDATE,
                        //To_Date = TODATE
                    });

                    streamWriter.Write(json);
                    streamWriter.Flush();
                    streamWriter.Close();
                }

                using (HttpWebResponse response = (HttpWebResponse)httpWebRequest.GetResponse())
                {

                    Stream Answer = response.GetResponseStream();
                    StreamReader _Answer = new StreamReader(Answer);
                    string page = _Answer.ReadToEnd();
                    response.Close();
                    object deserialize = JsonConvert.DeserializeObject(page);
                    root = JsonConvert.DeserializeObject<Root>(deserialize.ToString());
                }

                if (root.EmployeeCount > 0)
                {
                    Console.WriteLine("ZingToMIS: Total No Of Records : " + root.EmployeeCount.ToString());
                    WriteError("ZingToMIS:Total No Of Records : " + root.EmployeeCount.ToString());
                    mISDBModel = MAPZingModelToMISModel(root);

                    foreach (object item in root.employees)
                    {
                        if (((Employees)item).employeecode.ToString() == "4976")
                        {
                        }
                    }
                    WriteError("ZingToMIS:Start Push To DB: " + mISDBModel.Count.ToString());
                    Console.WriteLine("ZingToMIS: Start Push To DB:");

                   PushToDB(mISDBModel);

                    WriteError("ZingToMIS:END Push To DB: " + mISDBModel.Count.ToString());
                    Console.WriteLine("ZingToMIS: Data Inserted Successfully ");

                    Console.WriteLine("ZingToMIS: ExcelDTTOSendMail Start");
                    WriteError("ZingToMIS:ExcelDTTOSendMail Start");

                    ExcelDTTOSendMail(mISDBModel);

                    WriteError("ZingToMIS:ExcelDTTOSendMail End");
                    Console.WriteLine("ZingToMIS: ExcelDTTOSendMail End Successfully ");

                }
                else
                {
                    WriteError("ZingToMIS: No Records Found : ");
                    Console.WriteLine("ZingToMIS: No Records Found : ");
                }
            }
            catch (Exception ex)
            {
                WriteError("ZingToMIS: Error In Activity Start: " + ex.Message.ToString());
            }
        }
        public List<MISDBModel> MAPZingModelToMISModel(Root root)
        {
            List<MISDBModel> mISDBModel_List = new List<MISDBModel>();

            try
            {
                for (int i = 0; i < root.EmployeeCount; i++)
                {

                    MISDBModel mISDBModel = new MISDBModel();

                    mISDBModel.SALUTATION = string.IsNullOrWhiteSpace(root.employees[i].salutation.ToString()) ? "" : root.employees[i].salutation.ToString();
                    mISDBModel.EMPLOYEECODE = root.employees[i].employeecode.ToString();
                    mISDBModel.NAME = string.IsNullOrWhiteSpace(root.employees[i].employeename.ToString()) ? "" : root.employees[i].employeename.ToString().Replace("'", "");
                    mISDBModel.FIRST_NAME = string.IsNullOrWhiteSpace(root.employees[i].firstname.ToString()) ? "" : root.employees[i].firstname.ToString().Replace("'", "");
                    mISDBModel.LAST_NAME = string.IsNullOrWhiteSpace(root.employees[i].lastname.ToString()) ? "" : root.employees[i].lastname.ToString().Replace("'", "");
                    if (root.employees[i].attributes.Count > 0)
                    {
                        for (int x = 0; x < root.employees[i].attributes.Count - 1; x++)
                        {
                            if (root.employees[i].attributes[x].attribute_type_id == "69")
                                mISDBModel.DESIGNATION = string.IsNullOrWhiteSpace(root.employees[i].attributes[x].attribute_type_unit_desc.ToString()) ? " " : root.employees[i].attributes[x].attribute_type_unit_desc.ToString();
                            else if (root.employees[i].attributes[x].attribute_type_id == "41")
                                mISDBModel.DEPT_NAME = string.IsNullOrWhiteSpace(root.employees[i].attributes[x].attribute_type_unit_desc.ToString()) ? " " : root.employees[i].attributes[x].attribute_type_unit_desc.ToString();
                            else if (root.employees[i].attributes[x].attribute_type_id == "55")
                                mISDBModel.LOCATIONNAME = string.IsNullOrWhiteSpace(root.employees[i].attributes[x].attribute_type_unit_desc.ToString()) ? " " : root.employees[i].attributes[x].attribute_type_unit_desc.ToString();
                            else if (root.employees[i].attributes[x].attribute_type_id == "72")
                                mISDBModel.LOCATIONCODE = string.IsNullOrWhiteSpace(root.employees[i].attributes[x].attribute_type_unit_code.ToString()) ? " " : root.employees[i].attributes[x].attribute_type_unit_code.ToString();
                            else if (root.employees[i].attributes[x].attribute_type_id == "59")
                                mISDBModel.EMP_ROLE = string.IsNullOrWhiteSpace(root.employees[i].attributes[x].attribute_type_unit_desc.ToString()) ? "" : root.employees[i].attributes[x].attribute_type_unit_desc.ToString();
                        }
                    }
                    mISDBModel.DESIGNATION = String.IsNullOrEmpty(mISDBModel.DESIGNATION.ToString()) ? " " : mISDBModel.DESIGNATION.ToString();
                    mISDBModel.DEPT_NAME = String.IsNullOrEmpty(mISDBModel.DEPT_NAME.ToString()) ? " " : mISDBModel.DEPT_NAME.ToString();
                    mISDBModel.LOCATIONNAME = String.IsNullOrEmpty(mISDBModel.LOCATIONNAME.ToString()) ? " " : mISDBModel.LOCATIONNAME.ToString();
                    mISDBModel.LOCATIONCODE = String.IsNullOrEmpty(mISDBModel.LOCATIONCODE.ToString()) ? " " : mISDBModel.LOCATIONCODE.ToString();
                    mISDBModel.EMP_ROLE = String.IsNullOrEmpty(mISDBModel.EMP_ROLE.ToString()) ? " " : mISDBModel.EMP_ROLE.ToString();

                    mISDBModel.BIRTHDATE = Convert.ToDateTime(string.IsNullOrWhiteSpace(root.employees[i].dateofbirth.ToString()) ? "01-01-1900" : root.employees[i].dateofbirth.ToString());
                    mISDBModel.DATEOFJOINING = Convert.ToDateTime(string.IsNullOrWhiteSpace(root.employees[i].dateofjoining.ToString()) ? "01-01-1900" : root.employees[i].dateofjoining.ToString());
                    mISDBModel.DATEOFLEAVING = Convert.ToDateTime(string.IsNullOrWhiteSpace(root.employees[i].dateofleaving.ToString()) ? "01-01-1900" : root.employees[i].dateofleaving.ToString());
                    mISDBModel.EMPLOYEESTATUS = string.IsNullOrWhiteSpace(root.employees[i].employeestatus.ToString()) ? "" : root.employees[i].employeestatus.ToString();
                    mISDBModel.MOBILE = string.IsNullOrWhiteSpace(root.employees[i].mobile.ToString()) ? "" : root.employees[i].mobile.ToString();
                    mISDBModel.MOBILE_OFF = string.IsNullOrWhiteSpace(root.employees[i].alternatemobileno.ToString()) ? "" : root.employees[i].alternatemobileno.ToString();
                    mISDBModel.EMAIL_ID = string.IsNullOrWhiteSpace(root.employees[i].email.ToString()) ? "" : root.employees[i].email.ToString();
                    mISDBModel.SUPERVISORID = string.IsNullOrWhiteSpace(root.employees[i].reportingmanagercode.ToString()) ? "" : root.employees[i].reportingmanagercode.ToString();
                    mISDBModel.SUPERVISORNAME = string.IsNullOrWhiteSpace(root.employees[i].reportingmanagername.ToString()) ? "" : root.employees[i].reportingmanagername.ToString().Replace("'", ""); ;
                    mISDBModel.SUPERVISORMAILID = string.IsNullOrWhiteSpace(root.employees[i].reportingmanageremail.ToString()) ? "" : root.employees[i].reportingmanageremail.ToString();
                    if (root.employees[i].bankDetails.Count > 0)
                        mISDBModel.BANK_ACCOUNT_NO = string.IsNullOrWhiteSpace(root.employees[i].bankDetails[0].account_no.ToString()) ? "" : root.employees[i].bankDetails[0].account_no.ToString();
                    mISDBModel.LST_UPD_DT = Convert.ToDateTime(string.IsNullOrWhiteSpace(root.employees[i].created_date.ToString()) ? "01-01-1900" : root.employees[i].created_date.ToString());
                    mISDBModel.CREATED_DT = Convert.ToDateTime(string.IsNullOrWhiteSpace(root.employees[i].createddate.ToString()) ? "01-01-1900" : root.employees[i].createddate.ToString());
                    mISDBModel_List.Add(mISDBModel);
                }
                return mISDBModel_List;
            }
            catch (Exception ex)
            {
                WriteError("ZingToMIS: Error In Model Converting : " + ex.Message.ToString());
                mISDBModel_List = null;
                return mISDBModel_List;
            }
        }
        public void PushToDB(List<MISDBModel> mISDBModels)
        {
            try
            {
                using (OracleConnection conn = new OracleConnection(constr))
                {
                    //conn.Open();
                    foreach (MISDBModel mISDB in mISDBModels)
                    {
                        List<EMPCODE> EmpCodelist = new List<EMPCODE>();
                        string query1 = "SELECT DISTINCT EMPLOYEECODE FROM MISTEST.EMPLOYEE_LIST_FROM_ZING WHERE EMPLOYEECODE='" + mISDB.EMPLOYEECODE + "' ";
                        EmpCodelist = conn.Query<EMPCODE>(query1).ToList();
                        if (EmpCodelist.Count > 0)
                        {
                            string BackupQuery = " INSERT INTO MISTEST.EMPLOYEE_LIST_FROM_ZING_HIST(SALUTATION,EMPLOYEECODE,NAME,FIRST_NAME,LAST_NAME, " +
                                                " DESIGNATION,DEPT_NAME,BIRTHDATE,LOCATIONNAME,LOCATIONCODE,DATEOFJOINING,DATEOFLEAVING,EMPLOYEESTATUS, " +
                                                " MOBILE,MOBILE_OFF,EMAIL_ID,SUPERVISORID,SUPERVISORNAME,SUPERVISORMAILID,BANK_ACCOUNT_NO,LST_UPD_DT,CREATED_DT,HIST_DT,EMP_ROLE) " +
                                                " SELECT SALUTATION,EMPLOYEECODE,NAME,FIRST_NAME,LAST_NAME," +
                                                " DESIGNATION,DEPT_NAME,BIRTHDATE,LOCATIONNAME,LOCATIONCODE,DATEOFJOINING,DATEOFLEAVING,EMPLOYEESTATUS," +
                                                " MOBILE,MOBILE_OFF,EMAIL_ID,SUPERVISORID,SUPERVISORNAME,SUPERVISORMAILID,BANK_ACCOUNT_NO,LST_UPD_DT,CREATED_DT,SYSDATE,EMP_ROLE" +
                                                " FROM MISTEST.EMPLOYEE_LIST_FROM_ZING WHERE EMPLOYEECODE='" + mISDB.EMPLOYEECODE + "' ";
                            //" + "'12-SEP-2023'" + "

                            conn.Query(BackupQuery).FirstOrDefault();

                            string UpdateQuery = " UPDATE  MISTEST.EMPLOYEE_LIST_FROM_ZING SET " +
                                                 " SALUTATION='" + mISDB.SALUTATION.ToString().Trim() + "'," +
                                                 " NAME='" + mISDB.NAME.ToString().Trim() + "'," +
                                                 " FIRST_NAME='" + mISDB.FIRST_NAME.ToString().Trim() + "'," +
                                                 " LAST_NAME='" + mISDB.LAST_NAME.ToString().Trim() + "'," +
                                                 " DESIGNATION='" + mISDB.DESIGNATION.ToString().Trim() + "'," +
                                                 " DEPT_NAME='" + mISDB.DEPT_NAME.ToString().Trim() + "'," +
                                                 " BIRTHDATE= TO_DATE('" + mISDB.BIRTHDATE.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                                                 // " '" + mISDBModels[i].BIRTHDATE.ToString().Trim() + "'," +
                                                 " LOCATIONNAME='" + mISDB.LOCATIONNAME.ToString().Trim() + "'," +
                                                 " LOCATIONCODE='" + mISDB.LOCATIONCODE.ToString().Trim() + "'," +
                                                 " DATEOFJOINING=TO_DATE('" + mISDB.DATEOFJOINING.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                                                 // mISDBModels[i].DATEOFJOINING.ToString().Trim() + "," +
                                                 " DATEOFLEAVING=TO_DATE('" + mISDB.DATEOFLEAVING.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                                                 // mISDBModels[i].DATEOFLEAVING.ToString().Trim() + "," +
                                                 " EMPLOYEESTATUS='" + mISDB.EMPLOYEESTATUS.ToString().Trim() + "'," +
                                                 " MOBILE='" + mISDB.MOBILE.ToString().Trim() + "'," +
                                                 " MOBILE_OFF='" + mISDB.MOBILE_OFF.ToString().Trim() + "'," +
                                                 " EMAIL_ID='" + mISDB.EMAIL_ID.ToString().Trim() + "'," +
                                                 " SUPERVISORID='" + mISDB.SUPERVISORID.ToString().Trim() + "'," +
                                                 " SUPERVISORNAME='" + mISDB.SUPERVISORNAME.ToString().Trim() + "'," +
                                                 " SUPERVISORMAILID='" + mISDB.SUPERVISORMAILID.ToString().Trim() + "'," +
                                                 " BANK_ACCOUNT_NO='" + mISDB.BANK_ACCOUNT_NO.ToString().Trim() + "'," +
                                                 " LST_UPD_DT=TO_DATE('" + mISDB.LST_UPD_DT.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                                                 " CREATED_DT=TO_DATE('" + mISDB.CREATED_DT.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                                                 " EMP_ROLE='" + mISDB.EMP_ROLE.ToString().Trim() + "'" +
                                                 " WHERE EMPLOYEECODE='" + mISDB.EMPLOYEECODE + "' ";

                            conn.Query(UpdateQuery).FirstOrDefault();

                        }
                        else
                        {
                            string Insertquery = " INSERT INTO MISTEST.EMPLOYEE_LIST_FROM_ZING (SALUTATION,EMPLOYEECODE,NAME,FIRST_NAME,LAST_NAME," +
                                                  " DESIGNATION,DEPT_NAME,BIRTHDATE,LOCATIONNAME,LOCATIONCODE,DATEOFJOINING,DATEOFLEAVING,EMPLOYEESTATUS," +
                                                  " MOBILE,MOBILE_OFF,EMAIL_ID,SUPERVISORID,SUPERVISORNAME,SUPERVISORMAILID,BANK_ACCOUNT_NO,LST_UPD_DT,CREATED_DT,EMP_ROLE) " +
                                                  " VALUES ('" + mISDB.SALUTATION.ToString().Trim() + "'," +
                                                  " '" + mISDB.EMPLOYEECODE.ToString().Trim() + "'," +
                                                  " '" + mISDB.NAME.ToString().Trim() + "'," +
                                                  " '" + mISDB.FIRST_NAME.ToString().Trim() + "'," +
                                                  " '" + mISDB.LAST_NAME.ToString().Trim() + "'," +
                                                  " '" + mISDB.DESIGNATION.ToString().Trim() + "'," +
                                                  " '" + mISDB.DEPT_NAME.ToString().Trim() + "'," +
                                                   "TO_DATE('" + mISDB.BIRTHDATE.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                                                  // " '" + mISDBModels[i].BIRTHDATE.ToString().Trim() + "'," +
                                                  " '" + mISDB.LOCATIONNAME.ToString().Trim() + "'," +
                                                  "'" + mISDB.LOCATIONCODE.ToString().Trim() + "'," +
                                                  " TO_DATE('" + mISDB.DATEOFJOINING.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                                                  // mISDBModels[i].DATEOFJOINING.ToString().Trim() + "," +
                                                  " TO_DATE('" + mISDB.DATEOFLEAVING.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                                                  // mISDBModels[i].DATEOFLEAVING.ToString().Trim() + "," +
                                                  " '" + mISDB.EMPLOYEESTATUS.ToString().Trim() + "'," +
                                                  " '" + mISDB.MOBILE.ToString().Trim() + "'," +
                                                  " '" + mISDB.MOBILE_OFF.ToString().Trim() + "'," +
                                                  " '" + mISDB.EMAIL_ID.ToString().Trim() + "'," +
                                                  " '" + mISDB.SUPERVISORID.ToString().Trim() + "'," +
                                                  " '" + mISDB.SUPERVISORNAME.ToString().Trim() + "'," +
                                                  " '" + mISDB.SUPERVISORMAILID.ToString().Trim() + "'," +
                                                  " '" + mISDB.BANK_ACCOUNT_NO.ToString().Trim() + "'," +
                                                  "TO_DATE('" + mISDB.LST_UPD_DT.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                                                  "TO_DATE('" + mISDB.CREATED_DT.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                                                  " '" + mISDB.EMP_ROLE.ToString().Trim() + "')";

                            conn.Query(Insertquery).FirstOrDefault();
                        }
                    }
                }



                //using (OracleConnection conn = new OracleConnection(constr))
                //{
                //    conn.Open();

                //    for (int i = 0; i < mISDBModels.Count - 1; i++)
                //    {

                //        string query1 = " INSERT INTO MISTEST.EMPLOYEE_LIST_FROM_ZING (SALUTATION,EMPLOYEECODE,NAME,FIRST_NAME,LAST_NAME," +
                //                         " DESIGNATION,DEPT_NAME,BIRTHDATE,LOCATIONNAME,LOCATIONCODE,DATEOFJOINING,DATEOFLEAVING,EMPLOYEESTATUS," +
                //                         " MOBILE,MOBILE_OFF,EMAIL_ID,SUPERVISORID,SUPERVISORNAME,SUPERVISORMAILID,BANK_ACCOUNT_NO,LST_UPD_DT,CREATED_DT) " +
                //                         " VALUES ('" + mISDBModels[i].SALUTATION.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].EMPLOYEECODE.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].NAME.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].FIRST_NAME.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].LAST_NAME.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].DESIGNATION.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].DEPT_NAME.ToString().Trim() + "'," +
                //                          "TO_DATE('" + mISDBModels[i].BIRTHDATE.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                //                         // " '" + mISDBModels[i].BIRTHDATE.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].LOCATIONNAME.ToString().Trim() + "'," +                                    
                //                         "'" + mISDBModels[i].LOCATIONCODE.ToString().Trim()  + "'," +
                //                         " TO_DATE('" + mISDBModels[i].DATEOFJOINING.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                //                         // mISDBModels[i].DATEOFJOINING.ToString().Trim() + "," +
                //                         " TO_DATE('" +  mISDBModels[i].DATEOFLEAVING.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                //                         // mISDBModels[i].DATEOFLEAVING.ToString().Trim() + "," +
                //                         " '" + mISDBModels[i].EMPLOYEESTATUS.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].MOBILE.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].MOBILE_OFF.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].EMAIL_ID.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].SUPERVISORID.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].SUPERVISORNAME.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].SUPERVISORMAILID.ToString().Trim() + "'," +
                //                         " '" + mISDBModels[i].BANK_ACCOUNT_NO.ToString().Trim() + "'," +
                //                         "TO_DATE('" + mISDBModels[i].LST_UPD_DT.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss')," +
                //                         "TO_DATE('" + mISDBModels[i].CREATED_DT.ToString("dd-MMM-yyyy HH:mm:ss") + "','dd-mon-rrrr hh24:mi:ss'))";

                //        conn.Query(query1).FirstOrDefault();
                //    }

                //}
            }
            catch (Exception ex)
            {
                mISDBModels = null;
                WriteError("ZingToMIS: Error In Push To DB : " + ex.Message.ToString());
            }
        }
        public void WriteError(string errorMessage)
        {
            try
            {
                string path = strLogs + "\\ZinghrToMIS_" + DateTime.Today.ToString("ddMMyyyy") + ".txt";

                using (StreamWriter w = File.AppendText(path))
                {
                    string err = errorMessage + "\n";
                    w.WriteLine(err);
                    w.Flush();
                    w.Close();
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void ExcelDTTOSendMail(List<MISDBModel> mISDBModel)
        {
            try
            {
                string Filename = string.Empty;
                DataTable Data = ListToDataTable<MISDBModel>(mISDBModel);
                Filename = "ZingHrData_" + DateTime.Now.Date.AddDays(-1).ToString("dd-MMM-yyyy").ToUpper() + ".xlsx";
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage ExcelPkg = new ExcelPackage();
                ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
                using (ExcelRange Rng = wsSheet1.Cells[2, 2, 2, 2])
                {
                    Rng.Value = "ZingHr Data";
                    Rng.Style.Font.Size = 16;
                    Rng.Style.Font.Bold = true;
                }
                wsSheet1.Cells["A4"].LoadFromDataTable(Data, true);
                wsSheet1.Protection.IsProtected = false;
                wsSheet1.Protection.AllowSelectLockedCells = false;
                ExcelPkg.SaveAs(new FileInfo(FilePath + Filename));

                Console.WriteLine("ZingToMIS: Excel FIle Save Successfully ");
                try
                {
                    SendMail(Filename);
                    Console.WriteLine("ZingToMIS: Mail Send Successfully ");
                }
                catch (Exception ex)
                {
                    WriteError("ZingToMIS: Error In SendMail : " + ex.Message.ToString());
                }

            }
            catch (Exception ex)
            {
                WriteError("ZingToMIS: Error In ExcelDTTOSendMail : " + ex.Message.ToString());
            }

        }
        public DataTable ListToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);

            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {

                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {

                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        public void SendMail(string Filename)
        {
            try
            {
                string mailid = string.Empty;
                string strBody = string.Empty;
                string strFirstLine = string.Empty;
                string strSubject = string.Empty;

                mailid = "Cylsys.Mis@uti.co.in";

                strSubject = "ZingHr Data As Dated On " + DateTime.Now.Date.AddDays(-1).ToString("dd-MMM-yyyy").ToUpper();
                strBody = "<P>Respected Sir," +
                          "<P> Please Find the Attachement For ZingHr Data" +
                          "</BR></BR></BR>" +
                          "Thanks & Regards,</BR>" +
                          "Anil Kumar";

                MailMessage msgMail = new MailMessage();
                msgMail.Subject = strSubject;
                msgMail.Body = strBody;
                msgMail.From = new MailAddress(mailid.ToLower());

                string[] Multi = MAILTO.Trim().Split(',');
                foreach (string MultiEmailID in Multi)
                {
                    if (MultiEmailID.Contains("@"))
                    {
                        msgMail.To.Add(new MailAddress(MultiEmailID.ToLower()));
                    }
                }

                // msgMail.To.Add(new MailAddress(MAILTO.ToLower()));
                msgMail.Attachments.Add(new Attachment(FilePath + Filename));
                msgMail.IsBodyHtml = true;
                SmtpClient sc = new SmtpClient(strEmailIPAddress, 25);
                sc.Send(msgMail);
            }
            catch (Exception ex)
            {
                WriteError("ZingToMIS: Error In SendMail : " + ex.Message.ToString());
            }

        }
    }

}
public class Root
{
    public int Code { get; set; }
    public string Message { get; set; }
    public int EmployeeCount { get; set; }
    public List<Employees> employees { get; set; }
}
public class FamilyDetails
{
    public string first_name { get; set; }
    public string last_name { get; set; }
    public string middle_name { get; set; }
    public string date_of_birth { get; set; }
    public string age { get; set; }
    public string occupation { get; set; }
    public string is_dependant { get; set; }
    public string contact_no { get; set; }
    public string address { get; set; }
    public string family_relation_id { get; set; }
    public string family_relation_code { get; set; }
    public string family_relation_desc { get; set; }
    public string blood_group_id { get; set; }
    public string blood_group_code { get; set; }
    public string blood_group_desc { get; set; }
    public string qualification_id { get; set; }
    public string qualification_code { get; set; }
    public string qualification_desc { get; set; }
    public string stream_id { get; set; }
    public string stream_code { get; set; }
    public string stream_desc { get; set; }
    public string specialization_id { get; set; }
    public string sapecialization_code { get; set; }
    public string specialization_desc { get; set; }
    public string state_id { get; set; }
    public string state_code { get; set; }
    public string state_desc { get; set; }
    public string country_id { get; set; }
    public string country_code { get; set; }
    public string country_desc { get; set; }
    public string city_id { get; set; }
    public string city_code { get; set; }
    public string city_desc { get; set; }
    public long employeeid { get; set; }
}
public class Attributes
{
    public string attribute_type_id { get; set; } = string.Empty;
    public string attribute_type_desc { get; set; } = string.Empty;
    public string attribute_type_code { get; set; } = string.Empty;
    public string attribute_type_unit_id { get; set; } = string.Empty;
    public string attribute_type_unit_desc { get; set; } = string.Empty;
    public string attribute_type_unit_code { get; set; } = string.Empty;
    public long employeeid { get; set; }
}
public class BankDetails
{
    public string accounttype { get; set; }
    public string account_holder_name { get; set; }
    public string bank_name { get; set; }
    public string branch_name { get; set; }
    public string operation_type { get; set; }
    public string account_no { get; set; } = string.Empty;
    public string ifsc_code { get; set; }
    public long employeeid { get; set; }
}
public class Employees
{
    public string employeecode { get; set; }
    public string salutation { get; set; }
    public string employeename { get; set; }
    public string firstname { get; set; }
    public string lastname { get; set; }
    public string middlename { get; set; }
    public string fathername { get; set; }
    public string email { get; set; }
    public string gender { get; set; }
    public string mobile { get; set; }
    public string dateofbirth { get; set; }
    public string dateofjoining { get; set; }
    public string dateofconfirmation { get; set; }
    public string dateofleaving { get; set; }
    public string employeestatus { get; set; }
    public string age { get; set; }
    public string pfaccountnumber { get; set; }
    public string esicaccountnumber { get; set; }
    public string reportingmanagername { get; set; }
    public string reportingmanagercode { get; set; }
    public string lastmodified { get; set; }
    public string createddate { get; set; }
    public string empflag { get; set; }
    public string dateofresignation { get; set; }
    public string exitdate { get; set; }
    public string groupdoj { get; set; }
    public string exittypename { get; set; }
    public string exitreason1 { get; set; }
    public string exitreason2 { get; set; }
    public string domainid { get; set; }
    public string paymentdescription { get; set; }
    public string ptapplicable { get; set; }
    public string pfapplicable { get; set; }
    public string esicapplicable { get; set; }
    public string lwfapplicable { get; set; }
    public string pfdenotion { get; set; }
    public string employmenttype { get; set; }
    public string fnfprocessedmonth { get; set; }
    public string netpay { get; set; }
    public string fnfstatus { get; set; }
    public string recoverystatus { get; set; }
    public string ecodegeneratedby { get; set; }
    public string offerletterreferenceno { get; set; }
    public string attendancemanagerstatus { get; set; }
    public string nationality { get; set; }
    public string maritalstatus { get; set; }
    public string oldemployeecode { get; set; }
    public string attribute { get; set; }
    public string address { get; set; }
    public string alternatemobileno { get; set; }
    public string bloodgroup { get; set; }
    public string reportingmanageremail { get; set; }
    public string reportingmanagername2 { get; set; }
    public string reportingmanagercode2 { get; set; }
    public string reportingmanageremail2 { get; set; }
    public string personalemailaddress { get; set; }
    public long employeeid { get; set; }
    public DateTime created_date { get; set; }
    public string Department_name { get; set; }
    public List<FamilyDetails> familyDetails { get; set; }
    public List<Attributes> attributes { get; set; }
    public List<BankDetails> bankDetails { get; set; }
}
public class MISDBModel
{
    public string SALUTATION { get; set; } = string.Empty;
    public string EMPLOYEECODE { get; set; } = string.Empty;
    public string NAME { get; set; } = string.Empty;
    public string FIRST_NAME { get; set; } = string.Empty;
    public string LAST_NAME { get; set; } = string.Empty;
    public string DESIGNATION { get; set; } = string.Empty;
    public string DEPT_NAME { get; set; } = string.Empty;
    public DateTime BIRTHDATE { get; set; }
    public string LOCATIONNAME { get; set; } = string.Empty;
    public string LOCATIONCODE { get; set; } = string.Empty;
    public DateTime DATEOFJOINING { get; set; }
    public DateTime DATEOFLEAVING { get; set; }
    public string EMPLOYEESTATUS { get; set; } = string.Empty;
    public string MOBILE { get; set; } = string.Empty;
    public string MOBILE_OFF { get; set; } = string.Empty;
    public string EMAIL_ID { get; set; } = string.Empty;
    public string SUPERVISORID { get; set; } = string.Empty;
    public string SUPERVISORNAME { get; set; } = string.Empty;
    public string SUPERVISORMAILID { get; set; } = string.Empty;
    public string BANK_ACCOUNT_NO { get; set; } = string.Empty;
    public DateTime LST_UPD_DT { get; set; }
    public DateTime CREATED_DT { get; set; }
    public string EMP_ROLE { get; set; } = string.Empty;
}
public class EMPCODE
{
    public string EMPLOYEECODE { get; set; }
    // public DateTime LST_UPD_DT { get; set; }
}


