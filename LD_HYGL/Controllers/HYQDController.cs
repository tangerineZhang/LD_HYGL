using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using LD.DAL;
using System.Data;
using System.Configuration;

using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;

namespace LD_HYGL.Controllers
{
    public class HYQDController : Controller
    {
        public ActionResult Index()
        {
            PublicDao pPublicDao = new PublicDao();
            DataSet ds = new DataSet();

            ////生成二维码
            //QRCode pQRCode = new QRCode();
            //pQRCode.CreateQRCode("http://localhost:39398/HYQD/EndSignIn?mid=", "会后问卷");
            Session["LoginUN"] = "张伟东";
            Session["LoginUID"] = "zhangwdc";

            //生成会议列表
            //ds = pPublicDao.GetHYQDMeetingInfo("UserID = '" + Session["LoginUID"] + "' ", "m.BegTime desc,m.CreateDate desc");
            ds = pPublicDao.GetHYQDMeetingInfo("", "m.BegTime desc,m.CreateDate desc");
            ViewData["dsMeetingInfo"] = ds.Tables[0];

            return View();
        }


        public ActionResult SignInQRCode()
        {
            PublicDao pPublicDao = new PublicDao();
            DataSet ds = new DataSet();
            string mID = String.Empty;
            if (Request.QueryString["mID"] != null)
            {
                mID = Request.QueryString["mid"];
                ViewData["mID"] = mID;

                //获取会议信息
                ds = pPublicDao.GetHYQDMeetingInfo("ID = '" + mID + "' ", "");
                if (ds.Tables[0].Rows.Count > 0)
                {
                    ViewData["mName"] = ds.Tables[0].Rows[0]["Name"].ToString();
                    ViewData["mUser"] = ds.Tables[0].Rows[0]["UserName"].ToString();
                }
            }
            return View();
        }

        public ActionResult EndSignIn()
        {
            if (Request.QueryString["mid"] != null)
            {
                string mID = Request.QueryString["mid"];
                ViewData["mID"] = mID;
                Session["EndSignIn_mID"] = mID;
                return View();
            }
            else
            {
                return Redirect("Index");
            }

            //return View();
        }

        public ActionResult CreateMeeting()
        {
            //ViewData["LoginUN"] = Session["LoginUN"];
            return View();
        }

        public ActionResult BlankPage()
        {
            return View();
        }

        public ActionResult Test()
        {
            return View();
        }

        [HttpPost]
        public ActionResult AddMeeting(string Name, string BegTime, float Duration, string UserName)
        {
            try
            {
                int rCount = 0;
                string errorinfo = String.Empty;
                string guid = String.Empty;
                PublicDao pPublicDao = new PublicDao();
                bool isb = true;
                DataSet ds = new DataSet();
                //string UserID = Session["LoginUID"].ToString();
                string UserID = "";


                isb = pPublicDao.TransactionAddMeeting(Name, BegTime, Duration, UserName, UserID, ref guid, ref rCount, ref errorinfo);
                if (isb == true)
                {
                    //生成二维码
                    QRCode pQRCode = new QRCode();
                    pQRCode.CreateQRCode(ConfigurationManager.AppSettings["hyqbUrl"] + "/HYQD/EndSignIn?mid=" + guid, guid);

                    return Content("提交成功！");
                }
                else
                {
                    return Content(errorinfo);
                    //return Content("提交失败！");
                }
            }
            catch (Exception ex)
            {
                return Content(ex.Message.ToString());
            }
        }

        [HttpPost]
        public ActionResult AddEndSignIn(string UserName, string Q1, string Q2, string Q3, string Q4, string Q5)
        {
            int rCount = 0;
            string errorinfo = String.Empty;
            string guid = String.Empty;
            PublicDao pPublicDao = new PublicDao();
            bool isb = true;
            DataSet ds = new DataSet();

            string mID = String.Empty;
            if (Session["EndSignIn_mID"] != null)
            {
                mID = Session["EndSignIn_mID"].ToString();
            }
            isb = pPublicDao.TransactionAddEndSignIn(mID, UserName, Q1, Q2, Q3, Q4, Q5, ref rCount, ref errorinfo);
            if (isb == true)
            {
                //生成二维码
                QRCode pQRCode = new QRCode();
                pQRCode.CreateQRCode(ConfigurationManager.AppSettings["hyqbUrl"] + "/HYQD/EndSignIn?mid=" + guid, guid);

                return Content("提交成功！");
            }
            else
            {
                return Content(errorinfo);
                //return Content("提交失败！");
            }
        }

        [HttpPost]
        public ActionResult TestResponse()
        {
            return Content("测试成功！");
        }



        public ActionResult Excel(string IDs)
        {
            PublicDao dao = new PublicDao();
            try
            {
                IWorkbook book = new HSSFWorkbook();
                ISheet sheet = book.CreateSheet();

                IRow hr = sheet.CreateRow(0);
                hr.CreateCell(0).SetCellValue("会议名称");
                hr.CreateCell(1).SetCellValue("会议时间");
                hr.CreateCell(2).SetCellValue("主持人");
                hr.CreateCell(3).SetCellValue("问卷数量");
                hr.CreateCell(4).SetCellValue("问题1__A、会议不涉及决策（如选A，BCDE可跳过）;");
                hr.CreateCell(5).SetCellValue("问题1__B、会议决策点不明确;");
                hr.CreateCell(6).SetCellValue("问题1__C、决策点明确，但未做出决策;");
                hr.CreateCell(7).SetCellValue("问题1__D、通过会下沟通/报批/会签等方式即可决策，无需开会;");
                hr.CreateCell(8).SetCellValue("问题1__E、无问题;");
                hr.CreateCell(9).SetCellValue("问题2__A、材料冗长，问题暴露不充分，行动方案/资源需求不明确;");
                hr.CreateCell(10).SetCellValue("问题2__B、材料未提前24小时定稿、发送;");
                hr.CreateCell(11).SetCellValue("问题2__C、会前议题未交圈，材料未预审;");
                hr.CreateCell(12).SetCellValue("问题2__D、存在参会人不提前看材料，争论基础信息的现象;");
                hr.CreateCell(13).SetCellValue("问题2__E、无问题;");
                hr.CreateCell(14).SetCellValue("问题3__A、未提前沟通确认，参会范围、层级不合理;");
                hr.CreateCell(15).SetCellValue("问题3__B、同一部门（同一专业）超过1人参会;");
                hr.CreateCell(16).SetCellValue("问题3__C、议程安排不合理，有排队等待＞15min的情况;");
                hr.CreateCell(17).SetCellValue("问题3__D、无问题;");
                hr.CreateCell(18).SetCellValue("问题4__A、会议超时＞15min ;");
                hr.CreateCell(19).SetCellValue("问题4__B、会议时间长，效率低，有较大改进空间;");
                hr.CreateCell(20).SetCellValue("问题4__C、存在多平台重复汇报内容;");
                hr.CreateCell(21).SetCellValue("问题4__D、无问题;");
                hr.CreateCell(22).SetCellValue("问题5：其他突出问题（选填）：");
                IRow row1 = sheet.CreateRow(1);
                DataTable dt = dao.LdhySelectCount(IDs);
                if (dt.Rows.Count > 0)
                {
                    row1.CreateCell(0).SetCellValue(dt.Rows[0]["Name"].ToString());
                    row1.CreateCell(1).SetCellValue(dt.Rows[0]["BegTime"].ToString());
                    row1.CreateCell(2).SetCellValue(dt.Rows[0]["UserName"].ToString());
                    row1.CreateCell(3).SetCellValue(dt.Rows[0]["Num"].ToString());
                }

                DataRow[] Q1 = null;
                DataTable da = dao.LdhySelectData(IDs);
                #region 问题一
                Q1 = da.Select("Answer like '%A、%' and QuestionID='E0CD01F3-E65D-4DE5-B777-0F66C9B95BFB' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(4).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%B、%' and QuestionID='E0CD01F3-E65D-4DE5-B777-0F66C9B95BFB'");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(5).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%C、%' and QuestionID='E0CD01F3-E65D-4DE5-B777-0F66C9B95BFB' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(6).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%D、%' and QuestionID='E0CD01F3-E65D-4DE5-B777-0F66C9B95BFB'");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(7).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%E、%' and QuestionID='E0CD01F3-E65D-4DE5-B777-0F66C9B95BFB'");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(8).SetCellValue(Q1[0]["Countsum"].ToString());
                }

                #endregion
                #region 问题二


                Q1 = da.Select("Answer like '%A、%' and QuestionID='66B6A125-53A3-438E-B95A-C88730C300F9' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(9).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%B、%' and QuestionID='66B6A125-53A3-438E-B95A-C88730C300F9' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(10).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%C、%' and QuestionID='66B6A125-53A3-438E-B95A-C88730C300F9' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(11).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%D、%' and QuestionID='66B6A125-53A3-438E-B95A-C88730C300F9' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(12).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%E、%' and QuestionID='66B6A125-53A3-438E-B95A-C88730C300F9' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(13).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                #endregion
                #region 问题三
                Q1 = da.Select("Answer like '%A、%' and QuestionID='ED1D6BB4-CEF6-43B0-AFD1-3DED3B436525' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(14).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%B、%' and QuestionID='ED1D6BB4-CEF6-43B0-AFD1-3DED3B436525' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(15).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%C、%' and QuestionID='ED1D6BB4-CEF6-43B0-AFD1-3DED3B436525' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(16).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%D、%' and QuestionID='ED1D6BB4-CEF6-43B0-AFD1-3DED3B436525' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(17).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                #endregion
                #region 问题四
                Q1 = da.Select("Answer like '%A、%' and QuestionID='75B50C0F-FBD9-4CDC-8512-F05AEC3B5A16' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(18).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%B、%' and QuestionID='75B50C0F-FBD9-4CDC-8512-F05AEC3B5A16' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(19).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%C、%' and QuestionID='75B50C0F-FBD9-4CDC-8512-F05AEC3B5A16' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(20).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                Q1 = da.Select("Answer like '%D、%' and QuestionID='75B50C0F-FBD9-4CDC-8512-F05AEC3B5A16' ");
                if (Q1.Length > 0)
                {
                    row1.CreateCell(21).SetCellValue(Q1[0]["Countsum"].ToString());
                }
                #endregion
                #region 问题五
                string mag = "";
                DataTable ds = dao.LdhySelectWD(IDs);
                for (int j = 0; j < ds.Rows.Count; j++)
                {
                    mag += ds.Rows[j]["Answer"].ToString()+"；"; 
                }
                if (mag!="")
                {
                    row1.CreateCell(22).SetCellValue(mag);
                }
                #endregion

                MemoryStream memory = new MemoryStream();
                book.Write(memory);
                memory.Seek(0, SeekOrigin.Begin);

                return File(memory, "application/vnd.ms-execl", dt.Rows[0]["Name"].ToString() + "问卷"+DateTime.Now.ToLongDateString()+".xls");

            }
            catch (Exception ex)
            {
                return View(ex.ToString());
                // return Content(ex.Message.ToString());
            }
        }



    }


}
