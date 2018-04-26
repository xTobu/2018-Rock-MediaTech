using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http.Filters;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Web.Http.Controllers;
using System.Configuration;

namespace _2018_MediaTech
{
    public class JWTAuthorizationFilter : ActionFilterAttribute
    {
        public override void OnActionExecuting(HttpActionContext actionContext)
        {
            HttpStatusCode statusCode;
            string token;
            //determine whether a jwt exists or not
            if (!TryRetrieveToken(actionContext, out token))
            {
                statusCode = HttpStatusCode.Unauthorized;
                //allow requests with no token - whether a action method needs an authentication can be set with the claimsauthorization attribute
                setErrorResponse(50008, statusCode, actionContext, "Error Token");
                base.OnActionExecuting(actionContext);
                return;
            }

            try
            {
                string sec = ConfigurationManager.AppSettings["JWT_SecretKey"];
                var now = DateTime.UtcNow;
                var securityKey = new Microsoft.IdentityModel.Tokens.SymmetricSecurityKey(System.Text.Encoding.Default.GetBytes(sec));


                SecurityToken securityToken;
                JwtSecurityTokenHandler handler = new JwtSecurityTokenHandler();
                TokenValidationParameters validationParameters = new TokenValidationParameters()
                {
                    ValidIssuer = ConfigurationManager.AppSettings["JWT_issuer"],
                    ValidAudience = ConfigurationManager.AppSettings["JWT_audience"],
                    ValidateLifetime = true,
                    ValidateIssuerSigningKey = true,
                    LifetimeValidator = this.LifetimeValidator,
                    IssuerSigningKey = securityKey
                };
                //extract and assign the user of the jwt
                Thread.CurrentPrincipal = handler.ValidateToken(token, validationParameters, out securityToken);
                HttpContext.Current.User = handler.ValidateToken(token, validationParameters, out securityToken);

                base.OnActionExecuting(actionContext);
                return;
            }
            catch (SecurityTokenValidationException e)
            {
                statusCode = HttpStatusCode.Unauthorized;
                setErrorResponse(50008, statusCode, actionContext, "Error Token");
                base.OnActionExecuting(actionContext);
                return;
            }
            catch (Exception ex)
            {
                statusCode = HttpStatusCode.Unauthorized;
                setErrorResponse(50008, statusCode, actionContext, "Error Token");
                base.OnActionExecuting(actionContext);
                return;

            }
            //base.OnActionExecuting(actionContext);


        }

        private static void setErrorResponse(int code, HttpStatusCode HttpStatusCode, HttpActionContext actionContext, string message)
        {
            setErrorResponse_res res = new setErrorResponse_res();
            setErrorResponse_res_data res_data = new setErrorResponse_res_data();
            
            res_data.message = message;
            res.code = code;
            res.data = res_data;

            var response = actionContext.Request.CreateResponse(HttpStatusCode, res);
            actionContext.Response = response;
        }
        private class setErrorResponse_res
        {
            public int code { get; set; }
            public setErrorResponse_res_data data
            {
                get; set;
            }
        }
        private class setErrorResponse_res_data
        {
            public string message { get; set; }
        }
        private static bool TryRetrieveToken(HttpActionContext actionContext, out string token)
        {
            token = null;

            if (actionContext.Request.Headers.Authorization == null || actionContext.Request.Headers.Authorization.Scheme != "Bearer")
            {
                return false;
            }
            var bearerToken = actionContext.Request.Headers.Authorization.ToString();
            token = bearerToken.StartsWith("Bearer ") ? bearerToken.Substring(7) : bearerToken;
            return true;
        }

        public bool LifetimeValidator(DateTime? notBefore, DateTime? expires, SecurityToken securityToken, TokenValidationParameters validationParameters)
        {
            if (expires != null)
            {
                if (DateTime.UtcNow < expires) return true;
            }
            return false;
        }
    }
}