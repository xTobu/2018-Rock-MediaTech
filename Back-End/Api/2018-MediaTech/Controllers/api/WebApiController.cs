using Microsoft.Practices.EnterpriseLibrary.Data;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Web.Http;
using System.Web.Http.Cors;

namespace _2018_MediaTech.Controllers.api
{
    [RoutePrefix("api")]
    public class WebApiController : ApiController
    {
        Database DB = new DatabaseProviderFactory().Create("Rock_2018_ConnectionString");
        Models.WebApi WebApi = new Models.WebApi();

        #region Example

        [HttpGet]
        [Route("Student")]
        public IHttpActionResult Index()
        {
            ResponseModel _objResponseModel = new ResponseModel();
            List<Student> students = new List<Student>();
            students.Add(new Student
            {
                ID = 101,
                Name = "A",
                Email = "A@gmail.com",
                Address = "Taipei"
            });
            students.Add(new Student
            {
                ID = 102,
                Name = "B",
                Email = "B@gmail.com",
                Address = "Tainan"
            });
            _objResponseModel.Data = students;
            _objResponseModel.Status = true;

            return Ok(_objResponseModel);
            //400
            //return BadRequest(); 

        }
        public class Student
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public string Address { get; set; }
            public string Email { get; set; }
        }
        public class ResponseModel
        {
            public bool Status { set; get; }
            public object Data { set; get; }
        }
        #endregion

        #region Admin

        [HttpGet]
        [Route("Persons")]
        //[EnableCors(origins: "*", headers: "*", methods: "GET,OPTIONS")]
        public IHttpActionResult Persons(int limit, int page, string words)
        {
            Models.WebApi.GetModel.Persons_req req = new Models.WebApi.GetModel.Persons_req();
            Models.WebApi.GetModel.Persons_res res = new Models.WebApi.GetModel.Persons_res();

            req.limit = limit;
            req.page = page;
            req.words = words;
            try
            {
                DataSet ds = new DataSet();
                ds = WebApi.dsPersons(DB, req);
                DataTable dtPersons = ds.Tables[0];
                DataTable dtOrders = ds.Tables[1];
                DataTable dtTotalCount = ds.Tables[2];

                //新增Columns到dtPersons
                dtPersons.Columns.Add("Tickets", typeof(List<Models.WebApi.GetModel.Table_res_Tickets>));
                //dtPersons.Columns.Add("Tickets_TotalAmount", typeof(Int32));
                //dtPersons.Columns.Add("Tickets_TotalCount", typeof(Int32));

                int Tickets_TotalAmount;
                int Tickets_TotalCount;
                //輪循dtPersons每一個資料
                foreach (DataRow drPersons in dtPersons.Rows)
                {
                    Tickets_TotalAmount = 0;
                    Tickets_TotalCount = 0;

                    //建立一個LIST
                    List<Models.WebApi.GetModel.Table_res_Tickets> Tickets = new List<Models.WebApi.GetModel.Table_res_Tickets>();

                    //輪循dtOrders每一個資料
                    foreach (DataRow drOrders in dtOrders.Rows)
                    {
                        //如果guid相同 新增進LIST
                        if (drOrders["person_guid"].ToString() == drPersons["guid"].ToString())
                        {
                            Models.WebApi.GetModel.Table_res_Tickets ticket = new Models.WebApi.GetModel.Table_res_Tickets();
                            ticket.type = Convert.ToInt32(drOrders["ticket_type"]);
                            ticket.count = Convert.ToInt32(drOrders["ticket_count"]);
                            ticket.amount = Convert.ToInt32(drOrders["ticket_amount"]);
                            Tickets_TotalAmount += ticket.amount;
                            Tickets_TotalCount += ticket.count;
                            Tickets.Add(ticket);

                        }
                    }

                    //ADD LIST TO DATAROW
                    // 增加<List>到 Row
                    drPersons["Tickets"] = Tickets;
                    //drPersons["Tickets_TotalAmount"] = Tickets_TotalAmount;
                    //drPersons["Tickets_TotalCount"] = Tickets_TotalCount;      
                }

                //200
                res.code = 20000;
                res.data = dtPersons;
                res.total = Convert.ToInt32(dtTotalCount.Rows[0]["TotalCount"].ToString());

                return Ok(res);
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //500 伺服器錯誤
                return InternalServerError();
            }
        }



        [HttpPut]
        [Route("Update")]
        //[FromBody] 不太清楚  但不加的話在POST時必須有DATA
        //public IHttpActionResult Post([FromBody]RequestModel req)
        public IHttpActionResult UpdateRow([FromBody]Models.WebApi.PutModel.Update_req req)
        {

            Models.WebApi.PutModel.Update_res modelRes = new Models.WebApi.PutModel.Update_res();
            Models.WebApi.PutModel.Update_res_data modelRes_data = new Models.WebApi.PutModel.Update_res_data();
            try
            {
                WebApi.UpdateData(DB, req);

                modelRes.code = 20000;
                modelRes.data = modelRes_data;
                return Content(HttpStatusCode.OK, modelRes);

            }
            catch (Exception e)
            {
                modelRes.code = 400;
                modelRes_data.message = e.Message;

                modelRes.data = modelRes_data;
                return Content(HttpStatusCode.BadRequest, modelRes);

                //return BadRequest();
                //return InternalServerError();
            }


        }

        [HttpPut]
        [Route("SendEmail")]
        public IHttpActionResult SendEmailAndUpdate([FromBody]Models.WebApi.PutModel.SendEmail_req req)
        {
            Models.WebApi.PutModel.SendEmail_res res = new Models.WebApi.PutModel.SendEmail_res();
            Models.WebApi.PutModel.SendEmail_res_data res_data = new Models.WebApi.PutModel.SendEmail_res_data();
            res_data.guid = req.guid;
            try
            {
                //檢查GUID
                if (WebApi.Check_GUID(DB, req.guid))
                {
                    string ContactAddress = WebApi.Get_ContactAddress(DB, req.guid);
                    DataTable dt = WebApi.Get_PersonData(DB, req.guid);

                    WebApi.SendEmail(dt);
                    WebApi.Update_SendEmail(DB, req.guid);

                    //200 OK
                    res.code = 20000;
                    res.data = res_data;
                    return Content(HttpStatusCode.OK, res);
                }
                else
                {
                    //400 錯誤請求
                    res.code = 400;
                    res_data.message = "Bad Request - guid ";

                    res.data = res_data;
                    return Content(HttpStatusCode.BadRequest, res);
                }
            }
            catch (Exception e)
            {
                //400 錯誤請求
                res.code = 400;
                res_data.message = e.Message;

                res.data = res_data;
                return Content(HttpStatusCode.BadRequest, res);
            }
        }

        [HttpPut]
        [Route("SendEmailNotification")]
        public IHttpActionResult SendEmailNotificationAndUpdate([FromBody]Models.WebApi.PutModel.SendEmail_req req)
        {
            Models.WebApi.PutModel.SendEmail_res res = new Models.WebApi.PutModel.SendEmail_res();
            Models.WebApi.PutModel.SendEmail_res_data res_data = new Models.WebApi.PutModel.SendEmail_res_data();
            res_data.guid = req.guid;
            try
            {
                //檢查GUID
                if (WebApi.Check_GUID(DB, req.guid))
                {
                    string ContactAddress = WebApi.Get_ContactAddress(DB, req.guid);

                    WebApi.SendEmail_Notification(ContactAddress, req.guid);
                    WebApi.Update_SendEmail_Notification(DB, req.guid);

                    //200 OK
                    res.code = 20000;
                    res.data = res_data;
                    return Content(HttpStatusCode.OK, res);
                }
                else
                {
                    //400 錯誤請求
                    res.code = 400;
                    res_data.message = "Bad Request - guid ";

                    res.data = res_data;
                    return Content(HttpStatusCode.BadRequest, res);
                }
            }
            catch (Exception e)
            {
                //400 錯誤請求
                res.code = 400;
                res_data.message = e.Message;

                res.data = res_data;
                return Content(HttpStatusCode.BadRequest, res);
            }
        }

        [HttpPut]
        [Route("SendEmailRegitster")]
        public IHttpActionResult SendEmailRegitsterAndUpdate([FromBody]Models.WebApi.PutModel.SendEmail_req req)
        {
            Models.WebApi.PutModel.SendEmail_res res = new Models.WebApi.PutModel.SendEmail_res();
            Models.WebApi.PutModel.SendEmail_res_data res_data = new Models.WebApi.PutModel.SendEmail_res_data();
            res_data.guid = req.guid;
            try
            {
                //檢查GUID
                if (WebApi.Check_GUID(DB, req.guid))
                {
                    //string ContactAddress = WebApi.Get_ContactAddress(DB, req.guid);
                    DataTable dt = WebApi.Get_PersonData(DB, req.guid);

                    WebApi.SendEmail_Register(dt);

                    //200 OK
                    res.code = 20000;
                    res.data = res_data;
                    return Content(HttpStatusCode.OK, res);
                }
                else
                {
                    //400 錯誤請求
                    res.code = 400;
                    res_data.message = "Bad Request - guid ";

                    res.data = res_data;
                    return Content(HttpStatusCode.BadRequest, res);
                }
            }
            catch (Exception e)
            {
                //400 錯誤請求
                res.code = 400;
                res_data.message = e.Message;

                res.data = res_data;
                return Content(HttpStatusCode.BadRequest, res);
            }
        }

        #region JWT
        [HttpPost]
        [Route("JWT/Login")]
        public IHttpActionResult Authenticate([FromBody] LoginRequest login)
        {
            try
            {
                LoginResponse res = new LoginResponse();
                LoginResponse_data res_data = new LoginResponse_data();
                LoginRequest req = new LoginRequest();
                req.Username = login.Username.ToLower();
                req.Password = login.Password;


                bool isUsernamePasswordValid = false;

                if (login != null)
                    isUsernamePasswordValid = (req.Username == "admin" && req.Password == "webgene") ? true : false;
                // if credentials are valid
                if (isUsernamePasswordValid)
                {
                    res_data.token = createToken(req.Username);
                    res.code = 20000;
                    res.data = res_data;
                    return Ok(res);
                }
                else
                {
                    return BadRequest();
                }
            }
            catch (Exception ex)
            {
                return BadRequest();
            }
        }
        private string createToken(string username)
        {
            //Set issued at date
            DateTime issuedAt = DateTime.UtcNow;
            //set the time when it expires
            DateTime expires = DateTime.UtcNow.AddDays(7);
            expires = DateTime.UtcNow.AddHours(2);

            //http://stackoverflow.com/questions/18223868/how-to-encrypt-jwt-security-token
            var tokenHandler = new JwtSecurityTokenHandler();

            //create a identity and add claims to the user which we want to log in
            ClaimsIdentity claimsIdentity = new ClaimsIdentity(new[]
            {
                new Claim(ClaimTypes.Name, username)
            });

            string SecretKey = ConfigurationManager.AppSettings["JWT_SecretKey"];
            var now = DateTime.UtcNow;
            var securityKey = new Microsoft.IdentityModel.Tokens.SymmetricSecurityKey(System.Text.Encoding.Default.GetBytes(SecretKey));
            var signingCredentials = new Microsoft.IdentityModel.Tokens.SigningCredentials(securityKey, Microsoft.IdentityModel.Tokens.SecurityAlgorithms.HmacSha256Signature);


            //create the jwt
            var token =
                (JwtSecurityToken)
                    tokenHandler.CreateJwtSecurityToken(issuer: ConfigurationManager.AppSettings["JWT_issuer"], audience: ConfigurationManager.AppSettings["JWT_audience"],
                        subject: claimsIdentity, notBefore: issuedAt, expires: expires, signingCredentials: signingCredentials);
            var tokenString = tokenHandler.WriteToken(token);

            return tokenString;
        }
        public class LoginRequest
        {
            public string Username { get; set; }
            public string Password { get; set; }
        }
        public class LoginResponse
        {
            public int code { get; set; }
            public LoginResponse_data data { get; set; }

        }
        public class LoginResponse_data
        {
            public string token { get; set; }
        }

        [JWTAuthorizationFilter]
        [HttpGet]
        [Route("JWT/Authorization")]
        public IHttpActionResult JWTAuthorization()
        {
            //GET TOKEN
            string Authorization = Request.Headers.Authorization.ToString();

            JWTAuthorizationResponse res = new JWTAuthorizationResponse();
            JWTAuthorizationResponse_data res_data = new JWTAuthorizationResponse_data();


            res_data.name = "root";
            res_data.roles = "admin";
            res_data.avatar = Authorization;

            res.code = 20000;
            res.data = res_data;

            return Ok(res);
        }
        public class JWTAuthorizationResponse
        {
            public int code { get; set; }
            public JWTAuthorizationResponse_data data { get; set; }
        }
        public class JWTAuthorizationResponse_data
        {
            public string name { get; set; }
            public string roles { get; set; }
            public string avatar { get; set; }

        }


        [HttpPost]
        [Route("JWT/Logout")]
        public IHttpActionResult JWTLogout()
        {
            JWTLogout_res res = new JWTLogout_res();
            res.code = 20000;

            return Ok(res);
        }
        public class JWTLogout_res
        {
            public int code { get; set; }
        }

        #endregion

        #region Google recaptcha
        [HttpPost]
        [Route("Recaptcha/Verify")]
        public IHttpActionResult RecaptchaVerify([FromBody] RecaptchaVerifyRequest req)
        {

            HttpResponseMessage responseMsg = new HttpResponseMessage();

            if (WebApi.recaptcha_siteverify(req.response))
            {
                return Ok();
            }
            else
            {
                return Unauthorized();
            }

        }
        public class RecaptchaVerifyRequest
        {
            public string response { get; set; }
        }
        #endregion

        #region NPOI
        [HttpGet]
        [Route("DownloadExcel")]
        public HttpResponseMessage DownloadExcel()
        {
            Database DB = new DatabaseProviderFactory().Create("Rock_2018_ConnectionString");
            Models.WebApi WebApi = new Models.WebApi();
            Models.NPOI NPOI = new Models.NPOI();
            Models.WebApi.GetModel.Persons_req req = new Models.WebApi.GetModel.Persons_req();

            req.limit = 20;
            req.page = 1;

            DataSet ds = new DataSet();
            ds = WebApi.dsPersons(DB, req);
            DataTable dtPersons = ds.Tables[0];

            using (var stream = new MemoryStream())
            {
                HSSFWorkbook workbook = NPOI.NPOIGridviewToExcel(dtPersons);
                workbook.Write(stream);

                var result = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(stream.ToArray())
                };
                result.Content.Headers.ContentDisposition =
                    new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
                    {
                        FileName = "Orders" + ".xls"
                    };
                result.Content.Headers.ContentType =
                    new MediaTypeHeaderValue("application/octet-stream");

                return result;
            }

            // processing the stream.

        }
        #endregion

        #endregion

        #region Site
        [HttpPost]
        [Route("Submit")]
        //[FromBody] 不太清楚  但不加的話在POST時必須有DATA
        //public IHttpActionResult Post([FromBody]RequestModel req)
        public IHttpActionResult submit([FromBody]Models.WebApi.PostModel.Submit_req req)
        {

            Models.WebApi.PostModel.Submit_res modelRes = new Models.WebApi.PostModel.Submit_res();
            try
            {
                int status = WebApi.Submit_Insert(DB, req);

                modelRes.status = status;
                modelRes.message = "success";

                return Content(HttpStatusCode.OK, modelRes);

            }
            catch (Exception e)
            {
                modelRes.status = 400;
                modelRes.message = e.Message;
                modelRes.field = e.Data["Field"]?.ToString() ?? "";
                return Content(HttpStatusCode.BadRequest, modelRes);

                //return BadRequest();
                //return InternalServerError();
            }


        }
        #endregion



    }
}
