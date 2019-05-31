using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
namespace ClassCommon
{
    public class CookieTools
    {
        #region 添加/修改cookie信息
        /// <summary>
        /// 添加/修改cookie信息
        /// </summary>
        /// <param name="cookieName">cookie名字</param>
        /// <param name="db">数据</param>
        /// <param name="domain">作用域名</param>
        /// <param name="dt">过期时间</param>
        public static void AddCookie(string cookieName, string db, string domain, DateTime dt)
        {
            if (IsExistCookie(cookieName))
            {
                HttpCookie cookie = HttpContent.Request.Cookies[cookieName];
                cookie.Value = db;
                cookie.Domain = domain;
                cookie.Expires = dt;
                HttpContext.Current.Response.Cookies.Add(cookie);
            }
            else
            {
                HttpCookie cookie = new HttpCookie(cookieName);
                cookie.Value = db;
                cookie.Domain = domain;
                cookie.Expires = dt;
                HttpContext.Current.Response.Cookies.Add(cookie);
            }
        }
        #endregion

        #region 删除cookie信息
        public static void DelectCookie(string cookieName, string domain)
        {
            if (IsExistCookie(cookieName))
            {
                HttpCookie cookie = HttpContext.Current.Request.Cookies[cookieName];
                cookie.Expires = DateTime.Now.AddDays(-1);
                cookie.Domain = domain;
                HttpContext.Current.Response.Cookies.Add(cookie);
            }
        }
        #endregion

        #region 获取当前cookie信息
        public static string GetCookieValue(string cookieName)
        {
            if (IsExistCookie(cookieName))
            {
                HttpCookie cookie = HttpContext.Current.Request.Cookies[cookieName];
                return cookie.Value;
            }
            return "";
        }
        #endregion

        #region 是否存在cookie信息
        /// <summary>
        /// 是否存在cookie信息
        /// </summary>
        /// <param name="cookieName">cookie名字</param>
        public static bool IsExistCookie(string cookieName)
        {
            HttpCookie cookie = HttpContext.Current.Request.Cookies[cookieName];
            if (cookie != null)
                return true;
            else
                return false;
        }
        #endregion
    }
}
