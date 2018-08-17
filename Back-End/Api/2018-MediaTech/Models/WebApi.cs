using Microsoft.Practices.EnterpriseLibrary.Data;
using Newtonsoft.Json;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.Script.Serialization;

namespace _2018_MediaTech.Models
{
    public class WebApi
    {
        Common modelCommon = new Common();
        ConfigData modelConfigData = new ConfigData();
        public string UserDomainName = System.Environment.UserDomainName;

        //public Database DB = new DatabaseProviderFactory().Create("Rock_2018_ConnectionString");
        public Database DB = new DatabaseProviderFactory().Create("WebGeneDB_ConnectionString");


        //Get DataSet TableData
        #region Route Persons

        public DataSet dsPersons(Database DB, GetModel.Persons_req req)
        {

            try
            {
                DataSet ds = new DataSet();
                using (DbCommand se = DB.GetSqlStringCommand(@"SELECT 
	                                                             [id]
	                                                            ,[guid]
	                                                            ,[company_name]
	                                                            ,[company_taxid]
	                                                            ,[company_receipt]
	                                                            ,[contact_name]
	                                                            ,[contact_phone]
	                                                            ,[contact_ext]
	                                                            ,[contact_email]
	                                                            ,[contact_city]
	                                                            ,[contact_district]
	                                                            ,[contact_address]	                                                            
	                                                            ,[payment_method]
	                                                            ,[payment_account]
	                                                            ,[remarks]
	                                                            ,[tickets_count]
	                                                            ,[tickets_amount]
	                                                            ,[wasPaid]
	                                                            ,[wasSentLetter]
	                                                            ,[wasSentNotification]
	                                                            ,[wasSentRegister]
	                                                            ,[isDelete]
	                                                            ,[ip]
	                                                            ,convert(varchar, [CreatedTime], 120) as [CreatedTime]
                                                            FROM 
                                                            (
                                                                SELECT  ROW_NUMBER() OVER(ORDER BY [CreatedTime] desc) AS 'RowNumber', 
                                                                        *
                                                                FROM    [rock2018_db].[dbo].[MediaTech_Persons] 
                                                                 where   [isDelete]=0
																  AND( [company_name]  LIKE @words
																  OR   [contact_name]  LIKE @words
																  OR   [contact_email] LIKE @words
                                                                  OR   [contact_phone] LIKE @words)
                                                            ) AS data
                                                            WHERE data.RowNumber BETWEEN @start AND @end

                                                            SELECT 
	                                                              [ticket_type]
	                                                            , [ticket_count]
	                                                            , [ticket_amount]
	                                                            , [person_id]
	                                                            , [person_guid]
                                                            FROM [MediaTech_Orders]  
                                                            INNER JOIN (

                                                            SELECT RowNumber,guid
                                                            FROM 
                                                            (
                                                                SELECT  ROW_NUMBER() OVER(ORDER BY [CreatedTime] desc) AS 'RowNumber', 
                                                                        guid
                                                                FROM    [rock2018_db].[dbo].[MediaTech_Persons]    
                                                                 where   [isDelete]=0
																  AND( [company_name]  LIKE @words
																  OR   [contact_name]  LIKE @words
																  OR   [contact_email] LIKE @words
                                                                  OR   [contact_phone] LIKE @words) 
                                                            ) AS data
                                                            WHERE data.RowNumber BETWEEN @start AND @end

                                                            ) as data
                                                            ON data.guid = [MediaTech_Orders].person_guid

                                                            SELECT COUNT(*) as TotalCount FROM [rock2018_db].[dbo].[MediaTech_Persons]  WITH (NOLOCK) where [isDelete]=0
																  AND( [company_name]  LIKE @words
																  OR   [contact_name]  LIKE @words
																  OR   [contact_email] LIKE @words
                                                                  OR   [contact_phone] LIKE @words)"))
                {
                    var a = req.page;
                    var b = req.limit;
                    DB.AddInParameter(se, "@start", DbType.Int32, 1 + (b * (a - 1)));
                    DB.AddInParameter(se, "@end", DbType.Int32, a * b);
                    DB.AddInParameter(se, "@words", DbType.String, "%" + req.words + "%");
                    ds = DB.ExecuteDataSet(se);

                }
                //200
                return ds;
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw e;
            }
        }


        #endregion


        //新增資料
        #region Route Submit

        public int Submit_Insert(Database DB, PostModel.Submit_req req)
        {
            ///檢查資料與欄位
            CheckModel_Submit_req(req);

            if (!modelCommon.IsValidEmail(req.Contact.email))
            {

                throw modelCommon.getException("error： req.Contact.email is not allow", "Contact.email");
            }

            string guid = modelCommon.GetGUID;
            DataTable dt = new DataTable();

            //資料庫連線           
            using (DbConnection dbconn = DB.CreateConnection())
            {
                dbconn.Open();
                using (DbTransaction dbtrans = dbconn.BeginTransaction())
                {
                    using (DbCommand dbcmdInsertPerson = DB.GetSqlStringCommand(@"
                                                                    INSERT INTO [dbo].[MediaTech_Persons]
                                                                     ([guid]
                                                                    ,[company_name]
                                                                    ,[company_taxid]
                                                                    ,[company_receipt]
                                                                    ,[contact_name]
                                                                    ,[contact_phone]
                                                                    ,[contact_ext]
                                                                    ,[contact_email]
                                                                    ,[contact_city]
                                                                    ,[contact_district]
                                                                    ,[contact_address]
                                                                    ,[payment_method]
                                                                    ,[payment_account]
                                                                    ,[remarks]
                                                                    ,[wasPaid]
                                                                    ,[wasSentLetter]
                                                                    ,[wasSentNotification]
                                                                    ,[isDelete]
                                                                    ,[ip]
                                                                    ,[CreatedTime])
                                                                    VALUES
                                                                    (@guid
                                                                    ,@company_name
                                                                    ,@company_taxid
                                                                    ,@company_receipt
                                                                    ,@contact_name
                                                                    ,@contact_phone
                                                                    ,@contact_ext
                                                                    ,@contact_email
                                                                    ,@contact_city
                                                                    ,@contact_district
                                                                    ,@contact_address
                                                                    ,@payment_method
                                                                    ,@payment_account
                                                                    ,@remarks
                                                                    ,@wasPaid
                                                                    ,@wasSentLetter
                                                                    ,@wasSentNotification
                                                                    ,@isDelete
                                                                    ,@ip
                                                                    ,@CreatedTime);
                                                                    SELECT SCOPE_IDENTITY() as id,@guid as guid"))
                    {
                        try
                        {
                            #region dbcmdInsertPerson
                            DB.AddInParameter(dbcmdInsertPerson, "@guid", DbType.String, guid);

                            DB.AddInParameter(dbcmdInsertPerson, "@company_name", DbType.String, req.Company.name);
                            DB.AddInParameter(dbcmdInsertPerson, "@company_taxid", DbType.String, req.Company.taxid);
                            DB.AddInParameter(dbcmdInsertPerson, "@company_receipt", DbType.String, req.Company.receipt);

                            DB.AddInParameter(dbcmdInsertPerson, "@contact_name", DbType.String, req.Contact.name);
                            DB.AddInParameter(dbcmdInsertPerson, "@contact_phone", DbType.String, req.Contact.phone);
                            DB.AddInParameter(dbcmdInsertPerson, "@contact_ext", DbType.String, req.Contact.ext);
                            DB.AddInParameter(dbcmdInsertPerson, "@contact_email", DbType.String, req.Contact.email);
                            DB.AddInParameter(dbcmdInsertPerson, "@contact_city", DbType.String, req.Contact.city);
                            DB.AddInParameter(dbcmdInsertPerson, "@contact_district", DbType.String, req.Contact.district);
                            DB.AddInParameter(dbcmdInsertPerson, "@contact_address", DbType.String, req.Contact.address);

                            DB.AddInParameter(dbcmdInsertPerson, "@payment_method", DbType.String, req.payment_method);

                            DB.AddInParameter(dbcmdInsertPerson, "@payment_account", DbType.String, req.payment_method == "credit" ? null : req.payment_account);

                            DB.AddInParameter(dbcmdInsertPerson, "@remarks", DbType.String, req.remarks);

                            DB.AddInParameter(dbcmdInsertPerson, "@wasPaid", DbType.Boolean, false);
                            DB.AddInParameter(dbcmdInsertPerson, "@wasSentLetter", DbType.Boolean, false);
                            DB.AddInParameter(dbcmdInsertPerson, "@wasSentNotification", DbType.Boolean, false);
                            DB.AddInParameter(dbcmdInsertPerson, "@isDelete", DbType.Boolean, false);
                            DB.AddInParameter(dbcmdInsertPerson, "@ip", DbType.String, modelCommon.GetIP);
                            DB.AddInParameter(dbcmdInsertPerson, "@CreatedTime", DbType.DateTime, modelCommon.TWtime());
                            #endregion


                            //儲存Person 並取得 id 與 guid
                            dt = DB.ExecuteDataSet(dbcmdInsertPerson, dbtrans).Tables[0];

                            int person_id = Convert.ToInt32(dt.Rows[0]["id"]);
                            string person_guid = dt.Rows[0]["guid"].ToString();
                            int tickets_amount = 0;
                            int tickets_count = 0;

                            foreach (PostModel.Submit_req_Ticket Ticket in req.Tickets)
                            {
                                using (DbCommand dbcmdInsertOrder = DB.GetSqlStringCommand(@"
                                                                    INSERT INTO [dbo].[MediaTech_Orders]
                                                                      ([ticket_type]
                                                                      ,[ticket_count]
                                                                      ,[ticket_amount]
                                                                      ,[person_id]
                                                                      ,[person_guid]
                                                                      ,[CreatedTime])
                                                                    VALUES
                                                                    (@ticket_type
                                                                    ,@ticket_count
                                                                    ,@ticket_amount
                                                                    ,@person_id
                                                                    ,@person_guid
                                                                    ,@CreatedTime);"))
                                {
                                    int amount = modelConfigData.TicketType[Convert.ToInt32(Ticket.type)] * Convert.ToInt32(Ticket.count);

                                    #region dbcmdInsertOrder
                                    DB.AddInParameter(dbcmdInsertOrder, "@ticket_type", DbType.Int32, Ticket.type);
                                    DB.AddInParameter(dbcmdInsertOrder, "@ticket_count", DbType.Int32, Ticket.count);
                                    DB.AddInParameter(dbcmdInsertOrder, "@ticket_amount", DbType.Int32, amount);
                                    DB.AddInParameter(dbcmdInsertOrder, "@person_id", DbType.Int32, person_id);
                                    DB.AddInParameter(dbcmdInsertOrder, "@person_guid", DbType.String, person_guid);
                                    DB.AddInParameter(dbcmdInsertOrder, "@CreatedTime", DbType.DateTime, modelCommon.TWtime());
                                    #endregion

                                    tickets_amount += amount;
                                    tickets_count += Convert.ToInt32(Ticket.count);

                                    DB.ExecuteNonQuery(dbcmdInsertOrder, dbtrans);
                                }

                            }


                            using (DbCommand dbcmdUpdateAmount = DB.GetSqlStringCommand(@"
                                                                    UPDATE [dbo].[MediaTech_Persons]
                                                                       SET [tickets_amount] = @tickets_amount
                                                                          ,[tickets_count] = @tickets_count
                                                                     WHERE id=@id"))
                            {
                                #region dbcmdUpdataAmount

                                DB.AddInParameter(dbcmdUpdateAmount, "@tickets_amount", DbType.Int32, tickets_amount);
                                DB.AddInParameter(dbcmdUpdateAmount, "@tickets_count", DbType.Int32, tickets_count);
                                DB.AddInParameter(dbcmdUpdateAmount, "@id", DbType.Int32, person_id);

                                #endregion

                                DB.ExecuteNonQuery(dbcmdUpdateAmount, dbtrans);
                            }

                            //寄信與更新

                            using (DbCommand dbcmdUpdate = DB.GetSqlStringCommand(@"
                                                                    UPDATE [dbo].[MediaTech_Persons]
                                                                       SET [wasSentRegister] = @wasSentRegister
                                                                     WHERE guid=@guid"))
                            {
                                DB.AddInParameter(dbcmdUpdate, "@wasSentRegister", DbType.Boolean, true);
                                DB.AddInParameter(dbcmdUpdate, "@guid", DbType.String, guid);
                                DB.ExecuteNonQuery(dbcmdUpdate, dbtrans);
                            }


                            dbtrans.Commit();

                            //寄信
                            SendEmail_Register(Get_PersonData(DB, guid));

                            return 200;
                        }

                        catch (Exception ex)
                        {
                            dbtrans.Rollback();
                            throw ex;

                        }
                    }
                }
            }
        }

        public void CheckIsNullOrEmpty(string name, string val)
        {

            if (string.IsNullOrEmpty(val))
            {
                throw new Exception("error：[ " + name + " ] is null or empty");
            }
        }
        public void CheckModel_Submit_req(PostModel.Submit_req req)
        {
            if (!modelCommon.isFinish())
            {
                throw modelCommon.getException("error：event was finished! " + DateTime.Now.ToString("yyyyMMddHHmmss"), "");
            }

            //檢查資料
            if (req == null)
                throw modelCommon.getException("no request body", "");
            //客製判斷 [payment_method] [payment_account]
            if (req.payment_method != "remit" && req.payment_method != "credit")
            {
                throw modelCommon.getException("error： req.payment_method is not allow", "payment_method");
            }
            if (req.payment_method == "remit" && req.payment_account == null)
            {
                throw modelCommon.getException("error： req.payment_account can not be null", "payment_account");
            }
            if (req.payment_method == "credit" && req.payment_account != null)
            {
                throw modelCommon.getException("error： req.payment_account must be null", "payment_account");
            }

            //收據
            if (req.Company.receipt != "double" && req.Company.receipt != "triple")
            {
                throw modelCommon.getException("error：req.Company.receipt is not allow", "Company.receipt");
            }

            if (req.Tickets.Count != 3)
            {
                throw modelCommon.getException("error：req.Tickets is not allow", "Tickets");
            }

            if ((String.IsNullOrEmpty(req.Tickets[0].count.ToString()) || req.Tickets[0].count == 0) && (String.IsNullOrEmpty(req.Tickets[1].count.ToString()) || req.Tickets[1].count == 0) && (String.IsNullOrEmpty(req.Tickets[2].count.ToString()) || req.Tickets[2].count == 0))
            {
                throw modelCommon.getException("error：req.Tickets is not allow", "Tickets");
            }

            //判斷早鳥
            foreach (PostModel.Submit_req_Ticket ticket in req.Tickets)
            {
                if (!modelCommon.isEarlyBird() && ticket.type == 0 && ticket.count > 0)
                {
                    throw modelCommon.getException("error：req.Tickets is not allow", "Tickets");
                }
            }




            //輪循 req 內的所有 KEY 是否有東西
            foreach (var item in req.GetType().GetProperties())
            {
                var prop = item.GetValue(req, null);

                //客製判斷 payment_account
                if (item.Name == "payment_account")
                {
                    if (prop == null)
                    {
                        continue;
                    }

                    if (prop.ToString().Length != 5)
                    {
                        throw modelCommon.getException("error： " + "payment_account" + " is not equl 5", "payment_account");
                    }
                }
                //客製判斷 remarks
                if (item.Name == "remarks")
                {
                    if (prop == null)
                    {
                        continue;
                    }
                }


                //如果是NULL, throw
                if (prop == null)
                {
                    throw modelCommon.getException("error： " + item.Name + " is null", item.Name);
                }
                if (String.IsNullOrEmpty(prop.ToString()))
                    throw modelCommon.getException("error： " + item.Name + " is empty", item.Name);

                //var islist = IsList(prop);
                //var isdictionary = IsDictionary(prop);

                // 如果不是 List
                // 不曉得如何檢查Class obj
                if (prop.GetType().IsClass && !(prop is IEnumerable))
                {
                    foreach (var each in prop.GetType().GetProperties())
                    {
                        var key = each.Name;
                        var val = each.GetValue(prop, null);


                        if (val == null)
                        {
                            //客製
                            if (item.Name == "Company" && (key == "name" || key == "taxid"))
                            {
                                continue;
                            }
                            //客製
                            if (item.Name == "Contact" && (key == "ext"))
                            {
                                continue;
                            }
                            throw modelCommon.getException("error： " + item.Name + "." + key + " is null", item.Name + "." + key);
                        }
                        if (String.IsNullOrEmpty(val.ToString()))
                            throw modelCommon.getException("error： " + item.Name + "." + key + " is empty", item.Name + "." + key);
                    }
                }

                //如果是 List
                if (prop.GetType().IsGenericType && prop is IEnumerable)
                {
                    int? c = null;

                    foreach (var obj in prop as IEnumerable)
                    {
                        if (c == null) { c = 0; }

                        //EACH ARRAY'S OBJ
                        foreach (var each in obj.GetType().GetProperties())
                        {
                            var key = each.Name;
                            var val = each.GetValue(obj, null);
                            if (val == null)
                            {
                                throw modelCommon.getException("error： " + item.Name + "[" + c + "]." + key + " is null", item.Name + "[" + c + "]." + key);
                            }

                            if (String.IsNullOrEmpty(val.ToString()))
                                throw modelCommon.getException("error： " + item.Name + "[" + c + "]." + key + " is empty", item.Name + "[" + c + "]." + key);

                        }
                        c++;

                    }
                    if (c == null)
                    {
                        throw new Exception();
                        throw modelCommon.getException("error： " + item.Name + ".length is 0", "");

                    }
                }


            }



        }

        public bool IsList(object o)
        {
            if (o == null) return false;
            return o is IList &&
                   o.GetType().IsGenericType &&
                   o.GetType().GetGenericTypeDefinition().IsAssignableFrom(typeof(List<>));
        }
        public bool IsDictionary(object o)
        {
            if (o == null) return false;
            return o is IDictionary &&
                   o.GetType().IsGenericType &&
                   o.GetType().GetGenericTypeDefinition().IsAssignableFrom(typeof(Dictionary<,>));
        }

        public void SendEmail_Register(DataTable dt)
        {
            string contactName = dt.Rows[0]["contact_name"].ToString();
            string contactEmail = dt.Rows[0]["contact_email"].ToString();
            string type_early = dt.Rows[0]["type_early"].ToString();
            string type_normal = dt.Rows[0]["type_normal"].ToString();
            string type_student = dt.Rows[0]["type_student"].ToString();
            string tickets_amount = dt.Rows[0]["tickets_amount"].ToString();

            MailMessage mailMessage = new MailMessage("'Media Tech' <service@media-tech.com.tw>", contactEmail);
            mailMessage.Subject = "2018《Media Tech媒體科技大會》 報名成功通知信";
            mailMessage.IsBodyHtml = true;
            mailMessage.Body = MailBody_Register(contactName, type_early, type_normal, type_student, tickets_amount);//E-mail內容
            mailMessage.BodyEncoding = System.Text.Encoding.UTF8;//E-mail編碼

            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Host = "smtp.gmail.com";
            smtpClient.Port = 587;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = new NetworkCredential("service@media-tech.com.tw", "mt2018adm");
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.EnableSsl = true;
            smtpClient.Timeout = 12000;


            try
            {
                smtpClient.Send(mailMessage);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string MailBody_Register(string contactName, string type_early, string type_normal, string type_student, string tickets_amount)
        {
            string url = "https://mediatech2018.webgene.com.tw/edm/";

            //string early = "";
            //string normal = "";
            //string student = "";
            string body;
            body = @"
<html xmlns='http://www.w3.org/1999/xhtml'>

<head>
    <meta http-equiv='Content-Type' content='text/html; charset=UTF-8' />
    <title>報名成功</title>
    <meta name='viewport' content='width=device-width, initial-scale=1.0' />
</head>
<style>
body,
html {
    padding: 0px;
    margin: 0px;
}
</style>

<body>
    <table width='100%' border='0' cellspacing='0' cellpadding='0'>
        <tr>
            <td align='center' bgcolor='#e8e8e8' style='display:block;'>
                <table width='100%' border='0' cellspacing='0' cellpadding='0'>
                    <tr>
                        <td align='center'><img src='" + url + @"images/header.jpg' style='max-width: 700px !important; width: 100% !important; display: block;' /></td>
                    </tr>
                    <tr>
                        <td align='center'>
                            <table border='0' cellspacing='0' cellpadding='30' style='max-width: 700px !important; width: 100% !important;'>
                                <tr>
                                    <td bgcolor='#f8f8f8'>
                                        <h1 style='font-size: 1.8em; font-family:Arial, Helvetica, sans-serif;'>2018《Media Tech媒體科技大會》 報名成功通知信</h1>
                                        <p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>學員您好
                                            <br /> 感謝您的支持，您已報名成功，請記得於報名後5日內付款完成。
                                        </p>
                                        <p style='font-size: 1.1em; line-height: 2;  font-family:Arial, Helvetica, sans-serif; background-color: #e8e8e8; padding: 15px 25px;'>
                                            訂購人：" + contactName + @"
                                            <br> 票種數量：早鳥票 " + type_early + @" 張 / 一般票 " + type_normal + @" 張 / 學生票 " + type_student + @" 張
                                            <br> 總金額：
                                            <span style='color: #43a046;'><strong>" + Convert.ToInt32(tickets_amount).ToString("N0") + @"</strong></span>元<br>
                                            <hr/>
                                            <br> 銀行名稱 | 華南銀行 世貿分行
                                            <br> 銀行帳號 | 156-10-0500556
                                            <br> 收款帳戶地址 | 110台北市信義區信義路四段458號
                                            <br> SWIFT CODE | HNBKTWTP156
                                            <br> 收款人姓名 | 滾石文化股份有限公司
                                            <br> 收款人地址 | 台北市光復南路290巷1號6樓之1
                                            <br>
                                        </p>
                                       
                                        <p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>
                                            我們將於活動前10日寄出入場證，再請您多加留意。<br><span style='color: #43a046;'>若於7/6(五)前未收到入場證之學員，請立即與我們聯絡！</span>

                                        </p>
                                        <h2 style=' font-family:Arial, Helvetica, sans-serif;'>活動資訊：</h2>
                                        <p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>2018 媒體科技大會（Media Tech）
                                            <br> 日期：2018 年 7 月 10 日－7 月 11 日
                                            <br> 地點：台北國際會議中心 201 會議室
                                            <br> 地址：台北市信義區信義路五段 1 號 2 樓</p>
                                        <p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>
                                            如有任何問題請來信 <a href='mailto:vivian.hung@rock.com.tw' style='text-decoration: underline; color: #000;'>vivian.hung@rock.com.tw</a>，或電洽請於周一至周五上班時間撥打 02-2721-6121 分機 358，聯絡 Vivian( 洪小姐 )，大會保留議程變動之權利，最新消息將公布於大會官網 <a href='www.media-tech.com.tw' target='_blank' style='text-decoration:underline; color:#000;'>www.media-tech.com.tw</a>
                                        </p>
                                        <p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>                                           
                                            <br>敬祝 安好
                                        </p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td align='center'><img src='" + url + @"images/footer.jpg' style='max-width: 700px !important; width: 100% !important; display: block;' /></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>

</html>";
            return body;
        }


        #endregion

        //更新資料
        #region Route Update

        public void UpdateData(Database DB, PutModel.Update_req req)
        {
            if (req.payment_method == "credit")
                req.payment_account = null;
            // 檢查資料與欄位
            // CheckModel_Update_req(req);

            if (!modelCommon.IsValidEmail(req.contact_email))
            {
                throw new Exception("error： req.Contact.email is not allow");
            }

            DataTable dt = new DataTable();

            //資料庫連線           
            using (DbConnection dbconn = DB.CreateConnection())
            {
                dbconn.Open();
                using (DbTransaction dbtrans = dbconn.BeginTransaction())
                {
                    using (DbCommand dbcmdUpdatePerson = DB.GetSqlStringCommand(@"
                                                                    UPDATE [dbo].[MediaTech_Persons]
                                                                       SET [company_name] =  @company_name
                                                                          ,[company_taxid] = @company_taxid
                                                                          ,[company_receipt] = @company_receipt
                                                                          ,[contact_name] = @contact_name
                                                                          ,[contact_phone] = @contact_phone
                                                                          ,[contact_ext] = @contact_ext
                                                                          ,[contact_email] = @contact_email
                                                                          ,[contact_city] = @contact_city
                                                                          ,[contact_district] = @contact_district
                                                                          ,[contact_address] = @contact_address
                                                                          ,[payment_method] = @payment_method
                                                                          ,[payment_account] = @payment_account
                                                                          ,[remarks] = @remarks
                                                                          ,[tickets_amount] = @tickets_amount
                                                                          ,[tickets_count] = @tickets_count
                                                                          ,[UpdateTime] = @UpdateTime
                                                                     WHERE [guid] = @guid"))
                    {
                        try
                        {
                            #region dbcmdInsertPerson


                            DB.AddInParameter(dbcmdUpdatePerson, "@company_name", DbType.String, req.company_name);
                            DB.AddInParameter(dbcmdUpdatePerson, "@company_taxid", DbType.String, req.company_taxid);
                            DB.AddInParameter(dbcmdUpdatePerson, "@company_receipt", DbType.String, req.company_receipt);

                            DB.AddInParameter(dbcmdUpdatePerson, "@contact_name", DbType.String, req.contact_name);
                            DB.AddInParameter(dbcmdUpdatePerson, "@contact_phone", DbType.String, req.contact_phone);
                            DB.AddInParameter(dbcmdUpdatePerson, "@contact_ext", DbType.String, req.contact_ext);
                            DB.AddInParameter(dbcmdUpdatePerson, "@contact_email", DbType.String, req.contact_email);
                            DB.AddInParameter(dbcmdUpdatePerson, "@contact_city", DbType.String, req.contact_city);
                            DB.AddInParameter(dbcmdUpdatePerson, "@contact_district", DbType.String, req.contact_district);
                            DB.AddInParameter(dbcmdUpdatePerson, "@contact_address", DbType.String, req.contact_address);

                            DB.AddInParameter(dbcmdUpdatePerson, "@payment_method", DbType.String, req.payment_method);
                            DB.AddInParameter(dbcmdUpdatePerson, "@payment_account", DbType.String, req.payment_account);

                            DB.AddInParameter(dbcmdUpdatePerson, "@remarks", DbType.String, req.remarks);

                            DB.AddInParameter(dbcmdUpdatePerson, "@tickets_amount", DbType.Int32, req.tickets_amount);
                            DB.AddInParameter(dbcmdUpdatePerson, "@tickets_count", DbType.Int32, req.tickets_count);

                            DB.AddInParameter(dbcmdUpdatePerson, "@UpdateTime", DbType.DateTime, modelCommon.TWtime());
                            DB.AddInParameter(dbcmdUpdatePerson, "@guid", DbType.String, req.guid);
                            #endregion


                            //更新Person
                            DB.ExecuteNonQuery(dbcmdUpdatePerson, dbtrans);

                            foreach (PutModel.Update_req_ticket Ticket in req.Tickets)
                            {
                                using (DbCommand dbcmdUpdateOrder = DB.GetSqlStringCommand(@"
                                                                    UPDATE [dbo].[MediaTech_Orders]
                                                                       SET [ticket_count] = @ticket_count
                                                                          ,[ticket_amount] = @ticket_amount
                                                                          ,[UpdateTime] = @UpdateTime
                                                                     WHERE [person_guid] = @person_guid 
                                                                       and [person_id] = @person_id
                                                                       and [ticket_type] = @ticket_type"))
                                {
                                    int amount = modelConfigData.TicketType[Convert.ToInt32(Ticket.type)] * Convert.ToInt32(Ticket.count);

                                    #region dbcmdInsertOrder
                                    DB.AddInParameter(dbcmdUpdateOrder, "@ticket_count", DbType.Int32, Ticket.count);
                                    DB.AddInParameter(dbcmdUpdateOrder, "@ticket_amount", DbType.Int32, amount);
                                    DB.AddInParameter(dbcmdUpdateOrder, "@UpdateTime", DbType.DateTime, modelCommon.TWtime());

                                    DB.AddInParameter(dbcmdUpdateOrder, "@person_guid", DbType.String, req.guid);
                                    DB.AddInParameter(dbcmdUpdateOrder, "@person_id", DbType.Int32, req.id);
                                    DB.AddInParameter(dbcmdUpdateOrder, "@ticket_type", DbType.Int32, Ticket.type);


                                    #endregion

                                    DB.ExecuteNonQuery(dbcmdUpdateOrder, dbtrans);
                                }

                            }
                            dbtrans.Commit();
                        }

                        catch (Exception ex)
                        {
                            dbtrans.Rollback();
                            throw ex;

                        }
                    }
                }
            }
        }
        public void CheckModel_Update_req(PutModel.Update_req req)
        {
            //檢查資料
            if (req == null)
                throw new Exception("no request body");

            //客製判斷 [payment_method] [payment_account]
            if (req.payment_method != "remit" && req.payment_method != "credit")
            {
                throw new Exception("error： req.payment_method is not allow");
            }
            if (req.payment_method == "remit" && req.payment_account == null)
            {
                throw new Exception("error： req.payment_account can not be null");
            }
            if (req.payment_method == "credit" && req.payment_account != null)
            {
                throw new Exception("error： req.payment_account must be null");
            }

            //輪循 req 內的所有 KEY 是否有東西
            foreach (var item in req.GetType().GetProperties())
            {
                var prop = item.GetValue(req, null);



                //客製判斷 payment_account
                if (item.Name == "payment_account")
                {
                    if (prop == null)
                    {
                        continue;
                    }

                    if (prop.ToString().Length != 5)
                    {
                        throw new Exception("error： " + "payment_account" + " is not equl 5");
                    }


                }

                //如果是NULL, throw
                if (prop == null)
                {
                    throw new Exception("error： " + item.Name + " is null");
                }
                if (String.IsNullOrEmpty(prop.ToString()))
                    throw new Exception("error： " + item.Name + " is empty");

                //var islist = IsList(prop);
                //var isdictionary = IsDictionary(prop);

                // 如果不是 List
                // 不曉得如何檢查Class obj
                if (prop.GetType().IsClass && !(prop is IEnumerable))
                {
                    foreach (var each in prop.GetType().GetProperties())
                    {
                        var key = each.Name;
                        var val = each.GetValue(prop, null);
                        if (val == null)
                        {
                            throw new Exception("error： " + item.Name + "." + key + " is null");
                        }
                        if (String.IsNullOrEmpty(val.ToString()))
                            throw new Exception("error： " + item.Name + "." + key + " is empty");
                    }
                }

                //如果是 List
                if (prop.GetType().IsGenericType && prop is IEnumerable)
                {
                    int? c = null;

                    foreach (var obj in prop as IEnumerable)
                    {
                        if (c == null) { c = 0; }

                        //EACH ARRAY'S OBJ
                        foreach (var each in obj.GetType().GetProperties())
                        {
                            var key = each.Name;
                            var val = each.GetValue(obj, null);
                            if (val == null)
                            {
                                throw new Exception("error： " + item.Name + "[" + c + "]." + key + " is null");
                            }

                            if (String.IsNullOrEmpty(val.ToString()))
                                throw new Exception("error： " + item.Name + "[" + c + "]." + key + " is empty");

                        }
                        c++;

                    }
                    if (c == null)
                    {
                        throw new Exception("error： " + item.Name + ".length is 0");
                    }
                }


            }



        }

        #endregion

        //寄出 付款成功通知信
        #region Route SendEmail

        public void Update_SendEmail(Database DB, string guid)
        {
            //資料庫連線           
            using (DbConnection dbconn = DB.CreateConnection())
            {
                dbconn.Open();
                using (DbTransaction dbtrans = dbconn.BeginTransaction())
                {
                    using (DbCommand dbcmdUpdataAmount = DB.GetSqlStringCommand(@"
                                                                    UPDATE [dbo].[MediaTech_Persons]
                                                                       SET [wasPaid] = @wasPaid,
                                                                           [wasSentLetter] = @wasSentLetter
                                                                     WHERE guid=@guid"))
                    {
                        try
                        {
                            {
                                #region dbcmdUpdataAmount
                                DB.AddInParameter(dbcmdUpdataAmount, "@wasPaid", DbType.Boolean, true);
                                DB.AddInParameter(dbcmdUpdataAmount, "@wasSentLetter", DbType.Boolean, true);
                                DB.AddInParameter(dbcmdUpdataAmount, "@guid", DbType.String, guid);
                                #endregion

                                DB.ExecuteNonQuery(dbcmdUpdataAmount, dbtrans);
                            }
                            dbtrans.Commit();
                        }

                        catch (Exception ex)
                        {
                            dbtrans.Rollback();
                            throw ex;

                        }
                    }
                }
            }
        }
        public void SendEmail(DataTable dt)
        {
            string contactEmail = dt.Rows[0]["contact_email"].ToString();
            string ticketsAmount = dt.Rows[0]["tickets_amount"].ToString();
            MailMessage mailMessage = new MailMessage("'Media Tech' <service@media-tech.com.tw>", contactEmail);
            mailMessage.Subject = "2018《Media Tech媒體科技大會》 付款成功通知信";
            mailMessage.IsBodyHtml = true;
            mailMessage.Body = MailBody(ticketsAmount);//E-mail內容
            mailMessage.BodyEncoding = System.Text.Encoding.UTF8;//E-mail編碼

            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Host = "smtp.gmail.com";
            smtpClient.Port = 587;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = new NetworkCredential("service@media-tech.com.tw", "mt2018adm");
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.EnableSsl = true;
            smtpClient.Timeout = 12000;


            try
            {
                smtpClient.Send(mailMessage);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string MailBody(string ticketsAmount)
        {
            string body;
            string url = "https://mediatech2018.webgene.com.tw/edm/";
            body = @"<html xmlns='http://www.w3.org/1999/xhtml'>

<head>
    <meta http-equiv='Content-Type' content='text/html; charset=UTF-8' />
    <title>付款成功</title>
    <meta name='viewport' content='width=device-width, initial-scale=1.0' />
</head>
<style>
body,
html {
    padding: 0px;
    margin: 0px;
}
</style>

<body>
    <table width='100%' border='0' cellspacing='0' cellpadding='0'>
        <tr>
            <td align='center' bgcolor='#e8e8e8' style='display:block;'>
                <table width='100%' border='0' cellspacing='0' cellpadding='0'>
                    <tr>
                        <td align='center'><img src='" + url + @"images/header.jpg' style='max-width: 700px !important; width: 100% !important; display: block;' /></td>
                    </tr>
                    <tr>
                        <td align='center'>
                            <table border='0' cellspacing='0' cellpadding='30' style='max-width: 700px !important; width: 100% !important;'>
                                <tr>
                                    <td bgcolor='#f8f8f8'>
                                        <h1 style='font-size: 1.8em; font-family:Arial, Helvetica, sans-serif;'>2018《Media Tech媒體科技大會》 付款成功通知信</h1>
                                        <p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>
                                            學員您好
                                            <br>大會已確認收到您的報名費，共<span style='color: #43a046;'> <strong>" + Convert.ToInt32(ticketsAmount).ToString("N0") + @"</strong> </span>元！
                                            <br>我們將於活動前 10 日寄出入場證，再請您多加留意，謝謝！
                                            <br><span style='color: #43a046;'>若於7/6(五)前未收到入場證之學員，請立即與我們聯絡！</span></p>
                                        <h2 style=' font-family:Arial, Helvetica, sans-serif;'>活動資訊：</h2>
                                        <p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>2018 媒體科技大會（Media Tech）
                                        <br> 日期：2018 年 7 月 10 日－7 月 11 日
                                        <br> 地點：台北國際會議中心 201 會議室
                                        <br> 地址：台北市信義區信義路五段 1 號 2 樓</p>
                                        
                                        <p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>
                                            如有任何問題請來信 <a href='mailto:vivian.hung@rock.com.tw' style='text-decoration: underline; color: #000;'>vivian.hung@rock.com.tw</a>，或電洽請於周一至周五上班時間撥打 02-2721-6121 分機 358，聯絡 Vivian( 洪小姐 )，大會保留議程變動之權利，最新消息將公布於大會官網 <a href='www.media-tech.com.tw' target='_blank' style='text-decoration:underline; color:#000;'>www.media-tech.com.tw</a>
                                        </p>
                                        <p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>                                          
                                            <br>敬祝 安好
                                        </p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td align='center'><img src='" + url + @"images/footer.jpg' style='max-width: 700px !important; width: 100% !important; display: block;' /></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>

</html>";
            return body;
        }
        #endregion


        //寄出 報到通知信
        #region Route SendEmail_Notification
        public void Update_SendEmail_Notification(Database DB, string guid)
        {
            //資料庫連線           
            using (DbConnection dbconn = DB.CreateConnection())
            {
                dbconn.Open();
                using (DbTransaction dbtrans = dbconn.BeginTransaction())
                {
                    using (DbCommand dbcmdUpdataAmount = DB.GetSqlStringCommand(@"
                                                                    UPDATE [dbo].[MediaTech_Persons]
                                                                       SET [wasSentNotification] = @wasSentNotification
                                                                     WHERE guid=@guid"))
                    {
                        try
                        {
                            {
                                #region dbcmdUpdataAmount
                                DB.AddInParameter(dbcmdUpdataAmount, "@wasPaid", DbType.Boolean, true);
                                DB.AddInParameter(dbcmdUpdataAmount, "@wasSentNotification", DbType.Boolean, true);
                                DB.AddInParameter(dbcmdUpdataAmount, "@guid", DbType.String, guid);
                                #endregion

                                DB.ExecuteNonQuery(dbcmdUpdataAmount, dbtrans);
                            }
                            dbtrans.Commit();
                        }

                        catch (Exception ex)
                        {
                            dbtrans.Rollback();
                            throw ex;

                        }
                    }
                }
            }
        }
        public void SendEmail_Notification(string address, string guid)
        {
            MailMessage mailMessage = new MailMessage("'Media Tech' <service@media-tech.com.tw>", address);
            mailMessage.Subject = "2018《Media Tech媒體科技大會》 報到通知信";
            mailMessage.IsBodyHtml = true;
            mailMessage.Body = MailBody_Notification(guid);//E-mail內容
            mailMessage.BodyEncoding = System.Text.Encoding.UTF8;//E-mail編碼

            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Host = "smtp.gmail.com";
            smtpClient.Port = 587;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = new NetworkCredential("service@media-tech.com.tw", "mt2018adm");
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.EnableSsl = true;
            smtpClient.Timeout = 12000;


            try
            {
                smtpClient.Send(mailMessage);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string MailBody_Notification(string guid)
        {
            string url = "https://mediatech2018.webgene.com.tw/edm/";

            string body;
            body = @"<html xmlns='http://www.w3.org/1999/xhtml'>

<head>
    <meta http-equiv='Content-Type' content='text/html; charset=UTF-8' />
    <title>報到通知</title>
    <meta name='viewport' content='width=device-width, initial-scale=1.0' />
</head>
<style>
body,
html {
    padding: 0px;
    margin: 0px;
}
</style>

<body>
    <table width='100%' border='0' cellspacing='0' cellpadding='0'>
        <tr>
            <td align='center' bgcolor='#e8e8e8' style='display:block;'>
                <table width='100%' border='0' cellspacing='0' cellpadding='0'>
                    <tr>
                        <td align='center'><img src='" + url + @"images/header.jpg' style='max-width: 700px !important; width: 100% !important; display: block;' /></td>
                    </tr>
                    <tr>
                        <td align='center'>
                            <table border='0' cellspacing='0' cellpadding='30' style='max-width: 700px !important; width: 100% !important;'>
                                <tr>報名成功通知信
                                    <td bgcolor='#f8f8f8'>
                                        <h1 style='font-size: 1.8em; font-family:Arial, Helvetica, sans-serif;'>2018《Media Tech媒體科技大會》 報到通知信</h1>
                                        <p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>學員您好<br />
                                        再次感謝您報名2018《Media Tech媒體科技大會》 。</p><p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>大會將於7月10日（二）早上8點30分<span style='color: #43a046;'> <strong>憑入場證</strong> </span>進行報到，並領取活動資料袋，現領一份。<br>
（若您是公司企業負責報名窗口，請協助告知參加學員）<span style='color: #43a046;'><br />
若於7/6(五)前未收到入場證之學員，請立即與我們聯絡！</span></p>
                                        <h2 style=' font-family:Arial, Helvetica, sans-serif;'>活動資訊：</h2>
                                        <p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>2018 媒體科技大會（Media Tech）
                                        <br> 日期：2018 年 7 月 10 日－7 月 11 日
                                        <br> 地點：台北國際會議中心 201 會議室
                                        <br> 地址：台北市信義區信義路五段 1 號 2 樓</p>
                                        
                                        <h2 style=' font-family:Arial, Helvetica, sans-serif;'>重要提醒：</h2>
                                        <ul style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>
                                        <li><span style='color: #43a046;'>活動當日敬請記得攜帶入場證，本單位恕不補發。</span></li>
<li>本活動英語演講場次備有翻譯人員，提供逐步翻譯和Q&A翻譯的協助。</li>
<li>本活動不提供中餐，需請學員自理。</li>
<li>主辦單位保有大會所有活動變更之權利。</li>
<li>最新消息將公布於大會官網 <a href='www.media-tech.com.tw' target='_blank' style='text-decoration:underline; color:#000;'>www.media-tech.com.tw</a></li>
                                        </ul>
                                        <p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>
                                            如有任何問題請來信 <a href='mailto:vivian.hung@rock.com.tw' style='text-decoration: underline; color: #000;'>vivian.hung@rock.com.tw</a>，或電洽請於周一至周五上班時間撥打 02-2721-6121 分機 358，聯絡 Vivian( 洪小姐 )，大會保留議程變動之權利，最新消息將公布於大會官網 <a href='www.media-tech.com.tw' target='_blank' style='text-decoration:underline; color:#000;'>www.media-tech.com.tw</a>
                                        </p>
                                      <p style='font-size: 1.3em; line-height: 2;  font-family:Arial, Helvetica, sans-serif;'>                                            
                                            <br>敬祝 安好
                                        </p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td align='center'><img src='" + url + @"images/footer.jpg' style='max-width: 700px !important; width: 100% !important; display: block;' /></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>

</html>";
            return body;
        }
        #endregion

        #region google recaptcha

        public bool recaptcha_siteverify(string response)
        {
            JavaScriptSerializer json = new JavaScriptSerializer();

            recaptcha_siteverify_res res = new recaptcha_siteverify_res();
            string FeedRequestUrl = string.Concat("https://www.google.com/recaptcha/api/siteverify?secret=6LcL-FEUAAAAAJDF4SYVsq-UTI478zOTLNdmcls9&response=" + response);
            HttpWebRequest feedRequest = (HttpWebRequest)WebRequest.Create(FeedRequestUrl);
            feedRequest.Method = "GET";
            feedRequest.Accept = "application/json";
            feedRequest.ContentType = "application/json; charset=utf-8";


            try
            {
                using (var HttpWebResponse = feedRequest.GetResponse() as HttpWebResponse)
                {
                    if (feedRequest.HaveResponse && response != null)
                    {
                        using (var reader = new StreamReader(HttpWebResponse.GetResponseStream()))
                        {
                            res = JsonConvert.DeserializeObject<recaptcha_siteverify_res>(reader.ReadToEnd());
                        }
                    }
                }
            }
            catch (WebException wex)
            {
                if (wex.Response != null)
                {
                    using (var errorResponse = (HttpWebResponse)wex.Response)
                    {
                        using (var reader = new StreamReader(errorResponse.GetResponseStream()))
                        {
                            string error = reader.ReadToEnd();
                            //TODO: use JSON.net to parse this string and look at the error message
                        }
                    }
                }
            }

            return res.success;
        }
        public class recaptcha_siteverify_res
        {
            public bool success { get; set; }
            //public string challenge_ts { get; set; }
            //public string hostname { get; set; }
            //public List<string> errorcodes {get;set;}
        }
        #endregion


        #region RESTful Model
        //GET
        public class GetModel
        {
            public class Persons_req
            {
                public int page { get; set; }
                public int limit { get; set; }
                public string words { get; set; }
            }
            public class Persons_res
            {
                public int code { get; set; }
                public int total { get; set; }
                public DataTable data { get; set; }

            }
            public class Table_res
            {
                public string id { get; set; }
                public string guid { get; set; }
                public string contact_name { get; set; }
                public string contact_phone { get; set; }
                public string contact_email { get; set; }
                public string company_name { get; set; }
                public string company_taxid { get; set; }
                public string company_address { get; set; }
                public string payment_account { get; set; }
                List<Table_res_Tickets> tickets { get; set; }
                public string vegetarian { get; set; }
                public string non_vegetarian { get; set; }
                public string total_amount { get; set; }
                public string wasPaid { get; set; }
                public string wasSentLetter { get; set; }
                public string ip { get; set; }
                public string CreatedTime { get; set; }
            }
            public class Table_res_Tickets
            {
                public int type { get; set; }
                public int count { get; set; }
                public int amount { get; set; }
            }
        }

        //POST
        public class PostModel
        {
            public class Submit_req
            {

                public Submit_req_Company Company { get; set; }
                public Submit_req_Contact Contact { get; set; }
                public List<Submit_req_Ticket> Tickets { get; set; }
                public string remarks { get; set; }
                public string payment_method { get; set; }
                public string payment_account { get; set; }
            }
            public class Submit_req_Company
            {
                public string name { get; set; }
                public string taxid { get; set; }
                public string receipt { get; set; }
            }
            public class Submit_req_Contact
            {
                public string name { get; set; }
                public string phone { get; set; }
                public string ext { get; set; }
                public string email { get; set; }
                public string city { get; set; }
                public string district { get; set; }
                public string address { get; set; }
            }
            public class Submit_req_Ticket
            {
                public int type { get; set; }
                public int count { get; set; }
            }

            public class Submit_res
            {
                public int status { get; set; }
                public string message { get; set; }
                public string field { get; set; }
            }
        }

        //PUT
        public class PutModel
        {
            public class SendEmail_req
            {
                public string guid { get; set; }
                public string address { get; set; }
            }
            public class SendEmail_res
            {
                public int code { get; set; }
                public SendEmail_res_data data { get; set; }

            }
            public class SendEmail_res_data
            {
                public string guid { get; set; }
                public string message { get; set; }
            }

            public class Update_req
            {
                public int id { get; set; }
                public string guid { get; set; }

                public string company_name { get; set; }
                public string company_taxid { get; set; }
                public string company_receipt { get; set; }

                public string contact_name { get; set; }
                public string contact_phone { get; set; }
                public string contact_ext { get; set; }
                public string contact_email { get; set; }
                public string contact_city { get; set; }
                public string contact_district { get; set; }
                public string contact_address { get; set; }

                public string payment_method { get; set; }
                public string payment_account { get; set; }
                public string remarks { get; set; }
                public List<Update_req_ticket> Tickets { get; set; }
                public string tickets_count { get; set; }
                public string tickets_amount { get; set; }
            }
            public class Update_req_ticket
            {
                public int type { get; set; }
                public int count { get; set; }
            }
            public class Update_res
            {
                public int code { get; set; }
                public Update_res_data data { get; set; }

            }
            public class Update_res_data
            {
                public string guid { get; set; }
                public string message { get; set; }
            }
        }
        #endregion

        #region Common
        public bool Check_GUID(Database DB, string guid)
        {
            try
            {
                DataTable dt = new DataTable();
                using (DbCommand command = DB.GetSqlStringCommand(@"
                                                                SELECT COUNT(*) as count 
                                                                FROM [rock2018_db].[dbo].[MediaTech_Persons] 
                                                                WITH (NOLOCK)      
                                                                WHERE [guid] = @guid "))
                {
                    DB.AddInParameter(command, "@guid", DbType.String, guid);
                    dt = DB.ExecuteDataSet(command).Tables[0];
                    if (Convert.ToInt32(dt.Rows[0]["count"].ToString()) == 1)
                    {
                        return true;
                    }
                    return false;
                }
            }

            catch (Exception e)
            {
                throw e;
            }
        }

        public string Get_ContactAddress(Database DB, string guid)
        {
            try
            {
                DataTable dt = new DataTable();
                using (DbCommand command = DB.GetSqlStringCommand(@"SELECT [contact_email]  FROM [MediaTech_Persons]    
                                                                WHERE [guid] = @guid "))
                {
                    DB.AddInParameter(command, "@guid", DbType.String, guid);
                    dt = DB.ExecuteDataSet(command).Tables[0];
                    return dt.Rows[0]["contact_email"].ToString();

                }
            }

            catch (Exception e)
            {
                throw e;
            }
        }

        public DataTable Get_PersonData(Database DB, string guid)
        {
            try
            {
                DataTable dt = new DataTable();
                using (DbCommand command = DB.GetSqlStringCommand(@"                                                                with  

                                                                ty0 as  
                                                                (  
                                                                    SELECT ticket_count, person_guid
                                                                    FROM  [MediaTech_Orders]
                                                                    WHERE ticket_type = 0
                                                                ),
                                                                ty1 as  
                                                                (  
                                                                    SELECT ticket_count, person_guid
                                                                    FROM  [MediaTech_Orders]
                                                                    WHERE ticket_type = 1
                                                                ),  
                                                                ty2 as
                                                                (  
                                                                    SELECT ticket_count, person_guid
                                                                    FROM  [MediaTech_Orders]
                                                                    WHERE ticket_type = 2
                                                                ),
                                                                ty3 as
                                                                (  
                                                                   SELECT ticket_count, person_guid
                                                                    FROM  [MediaTech_Orders]
                                                                    WHERE ticket_type = 3
                                                                )



                                                                SELECT [contact_name]
                                                                      ,[contact_email]                                                                  
	                                                                  ,ISNULL (Type_0.ticket_count,0) as 'type_early'
	                                                                  ,ISNULL (Type_1.ticket_count,0) as 'type_normal'
	                                                                  ,ISNULL (Type_2.ticket_count,0) as 'type_student'	                                                                  
	                                                                  ,[tickets_amount] 
                                                                  FROM [rock2018_db].[dbo].[MediaTech_Persons] as a  
																 
  
                                                                  OUTER apply (
		                                                                SELECT ticket_count
                                                                        FROM ty0 c
                                                                        WHERE a.guid = c.person_guid)  Type_0
                                                                  OUTER apply (
		                                                                SELECT ticket_count
                                                                        FROM ty1 c
                                                                        WHERE a.guid = c.person_guid ) Type_1
                                                                  OUTER apply (
		                                                                SELECT ticket_count
                                                                        FROM ty2 c
                                                                        WHERE a.guid = c.person_guid ) Type_2

                                                                  where [isDelete] = 0 and [guid] = @guid"))
                {
                    DB.AddInParameter(command, "@guid", DbType.String, guid);
                    return DB.ExecuteDataSet(command).Tables[0];


                }
            }

            catch (Exception e)
            {
                throw e;
            }
        }
        #region Route Persons

        public DataTable dtSummaryReports(Database DB)
        {

            try
            {
                DataTable dt = new DataTable();
                using (DbCommand se = DB.GetSqlStringCommand(@"
                                                                DECLARE @priceEarlyBird int,
                                                                 @priceStudent int,
                                                                 @priceNormal int

                                                                SET @priceEarlyBird = 6000;
                                                                SET @priceNormal = 8000;
                                                                SET @priceStudent = 1000;

                                                                with  

                                                                ty0 as  
                                                                (  
                                                                    SELECT ticket_count, person_guid
                                                                    FROM  [MediaTech_Orders]
                                                                    WHERE ticket_type = 0
                                                                ),
                                                                ty1 as  
                                                                (  
                                                                    SELECT ticket_count, person_guid
                                                                    FROM  [MediaTech_Orders]
                                                                    WHERE ticket_type = 1
                                                                ),  
                                                                ty2 as
                                                                (  
                                                                    SELECT ticket_count, person_guid
                                                                    FROM  [MediaTech_Orders]
                                                                    WHERE ticket_type = 2
                                                                ),
                                                                ty3 as
                                                                (  
                                                                   SELECT ticket_count, person_guid
                                                                    FROM  [MediaTech_Orders]
                                                                    WHERE ticket_type = 3
                                                                )



                                                                SELECT [guid] as '訂單編號'
	                                                                  ,[company_name]  as '公司抬頭'
                                                                      ,[company_taxid] as '統一編號'
                                                                      ,[company_receipt] as '發票形式'

                                                                      ,[contact_name] as '姓名'
                                                                      ,[contact_phone] as '電話'
	                                                                  ,[contact_ext] as '分機'
                                                                      ,[contact_email] as '電子郵件'
	                                                                  ,[contact_city] as '縣市'
	                                                                  ,[contact_district] as '行政區'
	                                                                  ,[contact_address] as '完整街道位置'

                                                                      ,[payment_method] as '付款方式'
                                                                      ,[payment_account] as '匯款後三碼'
	                                                                  ,[remarks] as '備註'
	                                                                  ,CASE [wasPaid] WHEN 0 THEN '未付款' WHEN 1 THEN '已付款' END AS '是否付款'
                                                                      ,CASE [wasSentRegister] WHEN 0 THEN '未寄信' WHEN 1 THEN '已寄信' END AS '是否寄出報名成功通知信'
	                                                                  ,CASE [wasSentLetter] WHEN 0 THEN '未寄信' WHEN 1 THEN '已寄信' END AS '是否寄出付款成功通知信'
	                                                                  ,CASE [wasSentNotification] WHEN 0 THEN '未寄信' WHEN 1 THEN '已寄信' END AS '是否寄出報到通知信'
	                                                                  ,ISNULL (Type_0.ticket_count,0) as '早鳥票'
	                                                                  ,ISNULL (Type_1.ticket_count,0) as '一般票'
	                                                                  ,ISNULL (Type_2.ticket_count,0) as '學生票'
	                                                                  ,[tickets_count] as '總票數'
	                                                                  ,[tickets_amount] as '總票價'
	                                                                   ,[ip] as '來源'
                                                                      ,convert(varchar, [CreatedTime], 120) as '建立時間'
	                                                                  ,convert(varchar, [UpdateTime], 120) as '更新時間'
                                                                  FROM [rock2018_db].[dbo].[MediaTech_Persons] as a  
  
                                                                  OUTER apply (
		                                                                SELECT ticket_count
                                                                        FROM ty0 c
                                                                        WHERE a.guid = c.person_guid)  Type_0
                                                                  OUTER apply (
		                                                                SELECT ticket_count
                                                                        FROM ty1 c
                                                                        WHERE a.guid = c.person_guid ) Type_1
                                                                  OUTER apply (
		                                                                SELECT ticket_count
                                                                        FROM ty2 c
                                                                        WHERE a.guid = c.person_guid ) Type_2

                                                                  where [isDelete] = 0"))
                {
                    dt = DB.ExecuteDataSet(se).Tables[0];
                }
                //200
                return dt;
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw e;
            }
        }


        #endregion
        #endregion


    }
}