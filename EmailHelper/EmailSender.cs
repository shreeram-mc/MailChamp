using EmailHelper.Models;
using EmailHelper.Utilities;
using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Configuration;
using System.Net.Mail;

namespace EmailHelper
{
    public class EmailSender
    {
       public static  void SendEmail(Email email)
        {
            var service = ExchangeServiceUtil.GetExchangeService(email.EmailId, email.Password);

            try
            {                
                var serviceUrl = email.ExchangeUrl ?? "https://outlook.office365.com/ews/exchange.asmx";

                service.Url = new Uri(serviceUrl);

                EmailMessage emailMessage = new EmailMessage(service)
                {
                    Subject = email.Subject,
                    Body = new MessageBody(BodyType.HTML, email.Body)
                };

                if (email.ToRecipients.Contains(";"))
                {
                    var emails = email.ToRecipients.Split(';');

                    foreach(var item in emails)
                    {
                        if(!string.IsNullOrEmpty(item))
                            emailMessage.ToRecipients.Add(item.Trim());
                    }
                }
                else
                {
                    emailMessage.ToRecipients.Add(email.ToRecipients);
                } 

                emailMessage.Send();
            }
            catch (SmtpException exception)
            {
                string msg = $"Mail could not be sent (SmtpException):{exception.Message}";
               
                throw;
            }

            catch (AutodiscoverRemoteException exception)
            {
                string msg = $"Mail cannot be sent(AutodiscoverRemoteException): {exception.Message} ";
                
                throw;
            } 
        }
    }
}
