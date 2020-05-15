using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailHelper.Models
{
    public class Email
    {
        public string EmailId { get; set; }
        public string Password { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public List<EmailAddress> ToRecipients { get; set; }
        public string ExchangeUrl { get; set; }

        /// <summary>
        /// Either Token or Email/Pwd combination must be present
        /// </summary>
        public string Token { get; set; }
    }
}
