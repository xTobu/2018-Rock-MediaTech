using Microsoft.Practices.EnterpriseLibrary.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using _2018_MediaTech.Models;
namespace _2018_MediaTech.Controllers
{
    public class DefaultController : Controller
    {
        Models.Common modelCommon = new Models.Common();
        // GET: Default
        [HttpGet]
        [Route("")]
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        [Route("recaptcha")]
        public ActionResult recaptcha()
        {
            return View();
        }

        #region NPOI
        [HttpGet]
        [Route("NPOI/DownloadExcel")]
        public ActionResult NPOIDownloadExcel()
        {
            Database DB = new DatabaseProviderFactory().Create("Rock_2018_ConnectionString");
            Models.WebApi WebApi = new Models.WebApi();
            Models.NPOI NPOI = new Models.NPOI();
            //Models.WebApi.GetModel.Persons_req req = new Models.WebApi.GetModel.Persons_req();

            //req.limit = 20;
            //req.page = 1;

            //DataSet ds = new DataSet();
            //ds = WebApi.dsPersons(DB, req);
            //DataTable dtPersons = ds.Tables[0];
            DataTable dt = WebApi.dtSummaryReports(DB);
            using (var exportData = new MemoryStream())
            {
                HSSFWorkbook workbook = NPOI.NPOIGridviewToExcel(dt);
                workbook.Write(exportData);
                //string saveAsFileName = string.Format("Export-{0:d}.xls", DateTime.Now).Replace("/", "-");
                string saveAsFileName = "SummaryReport" + ".xls";
                byte[] bytes = exportData.ToArray();
                return File(bytes, "application/vnd.ms-excel", saveAsFileName);
            }
        }
        #endregion
    }
}