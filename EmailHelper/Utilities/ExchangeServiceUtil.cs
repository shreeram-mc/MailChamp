using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailHelper.Utilities
{
    public class ExchangeServiceUtil
    {
        /// <summary>
        /// Returns an exchange service (EWS)
        /// </summary>
        /// <param name="email">Email</param>
        /// <param name="password">Password</param>
        /// <param name="token">Token (optional)</param>
        /// <returns></returns>
        public static ExchangeService GetExchangeService(string email, string password, string token = "")
        {
            if (string.IsNullOrEmpty(token) && (string.IsNullOrEmpty(email) || string.IsNullOrEmpty(password)))
                throw new ArgumentNullException("Either token or Email/Password combination must be present ");

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
            {
                Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx")                
            };

            if (!string.IsNullOrEmpty(token))
            {                
                service.Credentials = new OAuthCredentials(token);
            }
            else
            {
                service.Credentials = new WebCredentials(email, password);
            }

            return service;
        }

    }
}
