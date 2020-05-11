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

        public static ExchangeService GetExchangeService(string email, string password)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1)
            {
                Credentials = new WebCredentials(email, password)
            };

            return service;
        }

    }
}
