using EmailHelper.Models;
using EmailHelper.Utilities;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;


namespace EmailHelper
{
    class Program
    {
        static void Main(string[] args)
        {
            string emailId, password;

            do
            {
                Console.WriteLine("Please enter your valid Office 365 Email ID");

                emailId = Console.ReadLine();

            } while (!Utils.IsValidEmail(emailId)); //Keep asking for a valid email ID if invalid

            do
            {
                Console.WriteLine("Please enter your Office 365 Password");

                password = ReadPassword();

            } while (string.IsNullOrEmpty(password)); //Keep asking for a password until entered

            ProcessOptions(emailId, password);

            Console.WriteLine("Please press c to start from main menu. Any other key to quit");

            if (Console.ReadLine() == "c")
            {
                Main(args);
            }
        }


        /// <summary>
        /// Ask the user with options to Send or Read Emails
        /// </summary>
        /// <param name="emailId"></param>
        /// <param name="password"></param>
        private static void ProcessOptions(string emailId, string password)
        {
            Console.WriteLine("\n1) Enter 1 to Compose Email\n2) Enter 2 to Read Top 100 Emails");

            var option = int.TryParse(Console.ReadLine(), out int val);

            if (!option)
            {
                Console.WriteLine("Invalid Option. Please try again! Press Ctrl+C to Quit");
                ProcessOptions(emailId, password);
            }

            switch (val)
            {
                case 1:
                    ComposeEmail(emailId, password);
                    break;

                case 2:
                    var emails = EmailReader.ReadEmails(new Email { EmailId = emailId, Password = password });
                    ProcessEmails(emails);
                    break;

                default:
                    ProcessOptions(emailId, password);
                    break;
            }
        }  

        private static void ComposeEmail(string emailId, string password)
        {
            var email = new Email
            {
                EmailId = emailId,
                Password = password
            };

            email.ToRecipients = GetAddresses();

            if (string.IsNullOrEmpty(email.ToRecipients))
            {
                Console.WriteLine("No valid recepients were found!");
                return;
            }

            Console.WriteLine("Enter the subject");
            email.Subject = Console.ReadLine() ?? "";

            Console.WriteLine("Please enter the message you want to send. Can include HTML tags also");

            email.Body = Console.ReadLine() ?? "";

            EmailSender.SendEmail(email);
        }

        private static string GetAddresses()
        {
            Console.WriteLine("Enter recepients emails address(s) as comma(,) seperated");

            var addresses = Console.ReadLine().Trim();

            if (string.IsNullOrEmpty(addresses.Trim()))
                GetAddresses();

            if (!addresses.Contains(",")) return addresses;

            var emails = addresses.Split(',');
            var validEmails = "";

            foreach (var emailAddress in emails)
            {
                if (Utils.IsValidEmail(emailAddress.Trim()))
                {
                    validEmails += emailAddress + ";";
                }
                else
                    Console.WriteLine($"Invalid: {emailAddress}. Ignored from the recepients list");
            }

            return validEmails;
        }


        private static void ProcessEmails(List<EmailMessage> emails)
        {
            foreach (var emailMsg in emails)
            {
                Console.WriteLine($"----------------------------------------Message ----------------------------------------------");
                Console.WriteLine($"On: {emailMsg.DateTimeReceived:dd/MM/yyyy HH:mm t}");
                Console.WriteLine($"From: {emailMsg.From.Address}");
                Console.WriteLine($"Subject: {emailMsg.Subject}");
                Console.WriteLine($"Message: {emailMsg.TextBody.Text}");
                Console.WriteLine($"----------------------------------------End of Message ----------------------------------------------\n\n");
            }
        }

        private static string ReadPassword()
        {
            string pass = "";
            do
            {
                ConsoleKeyInfo key = Console.ReadKey(true);

                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    pass += key.KeyChar;
                    Console.Write("*"); //Replace the pressed key with a *.
                }
                else
                {
                    //If Backspace was pressed, remove the last character from the 'pass' string and on the screen
                    if (key.Key == ConsoleKey.Backspace && pass.Length > 0)
                    {
                        pass = pass.Substring(0, (pass.Length - 1));
                        Console.Write("\b \b");
                    }
                    else if (key.Key == ConsoleKey.Enter)
                    {
                        break;
                    }
                }
            } while (true);

            return pass;
        }
    }
}
