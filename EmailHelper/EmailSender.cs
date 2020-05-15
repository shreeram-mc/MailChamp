using EmailHelper.Models;
using EmailHelper.Utilities;
using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Net.Mail;

namespace EmailHelper
{
    public class EmailSender
    {
        /// <summary>
        /// Send an Email with the given Email Account details
        /// </summary>
        /// <param name="email">Email</param>
        /// <exception cref="ArgumentNullException">Throw when Email Id or Password is Invalid</exception>
        /// <exception cref="SmtpException"></exception>
        /// <exception cref="AutodiscoverRemoteException"></exception>
        public static void SendEmail(Email email)
        {
            if (email == null)
                throw new ArgumentNullException();

            try
            {
                var service = ExchangeServiceUtil.GetExchangeService(email.EmailId, email.Password, email.Token);

                var serviceUrl = email.ExchangeUrl ?? "https://outlook.office365.com/ews/exchange.asmx";

                service.Url = new Uri(serviceUrl);

                EmailMessage emailMessage = new EmailMessage(service)
                {
                    Subject = email.Subject,
                    Body = new MessageBody(BodyType.HTML, email.Body),
                };

                emailMessage.ToRecipients.AddRange(email.ToRecipients);

                emailMessage.Send();
            }
            catch (SmtpException exception)
            {
                string msg = $"Mail could not be sent (SmtpException):{exception.Message}";
                //Log error message
                throw;
            }

            catch (AutodiscoverRemoteException exception)
            {
                string msg = $"Mail cannot be sent(AutodiscoverRemoteException): {exception.Message} ";
                //Log error message

                throw;
            }            
        }
    }
}
