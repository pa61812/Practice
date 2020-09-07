using Dapper;
using Microsoft.Ajax.Utilities;
using pratice.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using Member = pratice.Models.Member;
using pratice.Service;

namespace pratice.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            Animal animal = new Animal
            {
                AnimalID=1,
                AnimalName = "cat"
            };








            ////要連接資料庫所需要的 SqlConnection 物件
            //SqlConnection dataConnection = new SqlConnection();
            //string strMsgSelect = @"select*  from Member ";
            //Member member = new Member();
            ///*
            //string strMsgSelect2 = @"SELECT classDate, classStartTime FROM Arrangement 
            //WHERE teachID =@value1 ";
            //*/
            ////C# SQL Command 物件 SqlCommand( SQL語法 , SqlConnection )   
            //SqlCommand mySqlCmd = new SqlCommand(strMsgSelect, dataConnection);
            //try
            //{
            //    //設置資料庫位置
            //    //Initial Catalog = ArrangeSubjectDB   這是欲連接的資料庫
            //    //server = (Local)                               server的位置
            //    dataConnection.ConnectionString =WebConfigurationManager.ConnectionStrings["PracticeConnectionString"].ConnectionString;
            //    //連接資料庫
            //    dataConnection.Open();

            //    //塞C#參數到SQL語法中的參數
            //    //mySqlCmd.Parameters.AddWithValue("@value1",Account );

            //    //搜尋到的資料取出
            //    //連結讀取資料庫資料的元件            執行ExecuteReader()
            //    SqlDataReader dataReader = mySqlCmd.ExecuteReader();

            //    //開始讀資料
            //    while (dataReader.Read())
            //    {
            //        //dataReader ["欄位名稱"].ToString()    資料庫的資料
            //        //member.Account = dataReader["Account"].ToString();
            //        //member.Password = dataReader["Password"].ToString();
            //        //member.Name = dataReader["Name"].ToString();
            //        //member.Phone = dataReader["Phone"].ToString();
            //        //member.Tel = dataReader["Tel"].ToString();
            //        //member.Gender = dataReader["Gender"].ToString();
            //        //member.Birthday = Convert.ToDateTime(dataReader["Birthday"]);

            //        //foreach (var item in dataReader.AsQueryable())
            //        //{

            //        //}


           

            //    }

            //    //關閉讀取資料庫資料的元件
            //    dataReader.Close();

            //    //關閉資料庫
            //    dataConnection.Close();
            //}
            //catch (Exception e)
            //{
            //    string message = e.Message.ToString();
            //    return View(message);
            //}
            //finally
            //{
            //    //清除
            //    mySqlCmd.Cancel();
            //    dataConnection.Close();
            //    dataConnection.Dispose();
            //}



            return View(animal);
        }

        public ActionResult About()
        {  
            Member member = new Member();
            try
            {
                //要連接資料庫所需要的 SqlConnection 物件
                SqlConnection dataConnection = new SqlConnection();
                dataConnection.ConnectionString = WebConfigurationManager.ConnectionStrings["PracticeConnectionString"].ConnectionString;
                string strMsgSelect = "select*  from Member ";
                DataSet data = new DataSet();

                SqlDataAdapter da = new SqlDataAdapter(strMsgSelect, dataConnection);

                da.Fill(data);
                
                
               

                dataConnection.Close();

                return View();
            }
            catch (Exception e)
            {

                var Message = e.Message;
                return View(Message);
            }
          
           
        }

        public ActionResult Contact()
        {
            Member member = new Member();
            string strMsgSelect = "select*  from Member ";
            using (var connection = new SqlConnection(WebConfigurationManager.ConnectionStrings["PracticeConnectionString"].ConnectionString))
            {
                var anonymousList = connection.Query(strMsgSelect).ToList();
                var orderDetails = connection.Query<Member>(strMsgSelect).ToList();

                 return View(orderDetails);
            }


         
        }

        public ActionResult ExcelResult()
        {
          
                return View();
            
        }

        public JsonResult Import(HttpPostedFileBase file)
        {
            DataTable excelTable = new DataTable();
            var filename = file.FileName;
            //var filepath = Server.MapPath(string.Format("~/{0}","Files"));
            var filepath = Server.MapPath("~/Files");
            string path = Path.Combine(filepath, filename);
            file.SaveAs(path);

            excelTable = ImportExcel.GetExcelDataTable(path);

            List<Member> members = ImportExcel.DataTableToMember(excelTable);

            var result = Registers(members);

            ViewBag.result = result;

            return Json(result);

        }

        #region 新增MEMBER

        public string Registers(List<Member> members)
        {
            string Message = "成功";

            //要連接資料庫所需要的 SqlConnection 物件
            SqlConnection dataConnection = new SqlConnection();
            dataConnection.ConnectionString = WebConfigurationManager.ConnectionStrings["PracticeConnectionString"].ConnectionString;
            using (dataConnection)
            {
                dataConnection.Open();
                //加上BeginTrans
                using (var transaction = dataConnection.BeginTransaction())
                {
                    try
                    {
                        foreach (var item in members)
                        {

                            string strsql = "Insert into  Member " +
                                         "values(@Account,@Password,@Name,@Phone,@Tel,@Gender,@Birthday)" +
                                         " SELECT CAST(SCOPE_IDENTITY() as int)";
                            var id = dataConnection.Query<int>(strsql, item, transaction).SingleOrDefault();
                            //    dataConnection.Execute(strsql, member);
                            UserAccount userAccount = new UserAccount
                            {
                                ID = id,
                                Account = item.Account
                            };
                            InsertUserAccount(dataConnection, userAccount, transaction);

                            //正確就Commit
                            transaction.Commit();
                            dataConnection.Close();
                          
                        }
                    }
                    catch (Exception e)
                    {
                        dataConnection.Close();
                        Message = e.Message;
                       
                    }
                }
            }
            return Message;
        }
     
            //return Json(Message);


        public virtual JsonResult Registered(string UserID, string Pass, string CName, string Phone, string Tel, string Gender, string Birth)
        {


            string Message = "成功";

            Member member = new Member
            {
                Account = UserID,
                Password = Pass,
                Name = CName,
                Phone = Phone,
                Tel = Tel,
                Gender = Gender,
                Birthday = Convert.ToDateTime(Birth)
            };       
             //要連接資料庫所需要的 SqlConnection 物件
             SqlConnection dataConnection = new SqlConnection();
             dataConnection.ConnectionString = WebConfigurationManager.ConnectionStrings["PracticeConnectionString"].ConnectionString;
              using (dataConnection)
              {
                    dataConnection.Open();
                    //加上BeginTrans
                    using (var transaction = dataConnection.BeginTransaction())
                    {

                        try
                        {
                            string strsql = "Insert into  Member " +
                                         "values(@Account,@Password,@Name,@Phone,@Tel,@Gender,@Birthday)" +
                                         " SELECT CAST(SCOPE_IDENTITY() as int)";
                            var id = dataConnection.Query<int>(strsql, member,transaction).SingleOrDefault();
                            //    dataConnection.Execute(strsql, member);
                            UserAccount userAccount = new UserAccount
                            {
                                ID = id,
                                Account = UserID
                            };
                            InsertUserAccount(dataConnection, userAccount, transaction);

                            //正確就Commit
                            transaction.Commit();
                            dataConnection.Close();
                            return Json(Message);
                        }
                        catch (Exception e)
                        {
                        dataConnection.Close();
                        Message = e.Message;
                            return Json(Message);
                        }


                    }
              }   
            //return Json(Message);
        }
        public static void InsertUserAccount(IDbConnection conn, UserAccount userAccount, SqlTransaction transaction)
        {

            UserStaus userStaus = new UserStaus
            {
                ID = userAccount.ID,
                Account = userAccount.Account,
                ISUse = "Y",
                SDate = DateTime.Now
            };

            string strsql = "Insert into  UserStaus " +
                            "values(@ID,@Account,@SDate,@EDate,@ISUse)";



            conn.Execute(strsql, userStaus, transaction);
        }
        #endregion

    }
}