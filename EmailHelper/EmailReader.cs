using EmailHelper.Models;
using EmailHelper.Utilities;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Net;

namespace EmailHelper
{
    public class EmailReader
    {
        /// <summary>
        /// Read emails from the EWS
        /// </summary>
        /// <param name="email">email</param>
        /// <returns>All Emails from Inbox as a List of Messages</returns>
        /// <exception cref="ArgumentNullException"></exception>        
        public static List<EmailMessage> ReadEmails(Email email)
        {
            if (email == null || string.IsNullOrEmpty(email.EmailId) || string.IsNullOrEmpty(email.Password))
                throw new ArgumentNullException();

            var service = ExchangeServiceUtil.GetExchangeService(email.EmailId, email.Password);

            try
            {
                service.Url = new Uri(email.ExchangeUrl ?? "https://outlook.office365.com/ews/exchange.asmx");
                service.TraceEnabled = false;

                //By Default, exchange won't return all the properties of an email. We are specifying here to get Email Body also.
                //Body is in HTML format. TextBody contains the actual text message
                PropertySet propSet = new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.TextBody, ItemSchema.Body);

                //Setting a max emails to read per request. There might be a threshold on the number of emails to fetch in a request.
                ItemView view = new ItemView(1000);

                FindItemsResults<Item> foundItems = null;

                var list = new List<EmailMessage>();

                do
                {
                    //Get actual items from EWS inbox. 
                    foundItems = service.FindItems(WellKnownFolderName.Inbox, view);

                    foreach (EmailMessage emailMsg in foundItems)
                    {
                        emailMsg.Load(propSet); //Load those extra properties like body and body text

                        list.Add(emailMsg);
                    }

                    //Set the offset for next page results
                    view.Offset += foundItems.Items.Count;
                }
                while (foundItems.MoreAvailable == true); //If more items are there, coninue the loop.

                return list;
            }
            catch
            {
                //Log the error

                throw;
            }
        }
    }
}
