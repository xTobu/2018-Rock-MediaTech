using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _2018_MediaTech.Models
{
    public class Common
    {
        //取得 NewGuid
        public string NewGuid(int count)
        {
            return Guid.NewGuid().ToString("N").Substring(0, count);

        }

        //取得資料庫用的時間 字串化
        public DateTime TWtime()
        { //以台北時間為基準
            DateTime today = DateTime.Now;
            TimeZoneInfo est = TimeZoneInfo.FindSystemTimeZoneById("Taipei Standard Time"); // 取得台北時區與格林威治標準時間差
            DateTime TWTime = TimeZoneInfo.ConvertTime(today, est); // 轉換為台北時間
            //string time = TWTime.AddDays(0).ToString("yyyy-MM-dd HH:mm:ss");
            return TWTime.AddDays(0);
        }

        //取得GUID並將中間的 `-` 去掉
        public string GetGUID = Guid.NewGuid().ToString("N");

        public string GetIP = HttpContext.Current.Request.UserHostAddress;

        //判斷Email
        public bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        public Exception getException(string message, string field)
        {
            Exception e = new Exception(message);
            e.Data["Field"] = field;
            return e;
        }
        public bool isEarlyBird()
        {
            bool res;
            DateTime today = TWtime();
            DateTime EarlyBird = new DateTime(2018, 6, 9, 0, 0, 0);
            int result = DateTime.Compare(today, EarlyBird);
            if(result < 0)
            {
                //小於早鳥期限
                res = true;
            }
            else
            {
                //大於早鳥期限
                res = false;
            }
            return res;
        }

        public bool isFinish()
        {
            bool res;
            DateTime today = TWtime();
            DateTime Finish = new DateTime(2018, 7, 4, 23, 59, 59);
            int result = DateTime.Compare(today, Finish);
            if (result < 0)
            {
                //小於結束期限
                res = true;
            }
            else
            {
                //大於結束期限
                res = false;
            }
            return res;
        }
    }

}