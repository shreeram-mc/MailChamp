using EmailHelper;
using EmailHelper.Models;
using EmailHelper.Utilities;
using MailChampWin.Models;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;

namespace active_directory_wpf_msgraph_v2
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Set the scope for API call
        string[] scopes = new string[] { "https://outlook.office.com/EWS.AccessAsUser.All" };
        AuthenticationResult authResult = null;

        public MainWindow()
        {
            InitializeComponent();

            grdInbox.Visibility = Visibility.Hidden;
            grdCompose.Visibility = Visibility.Hidden;

            if (authResult != null)
                SignOutButton.Visibility = Visibility.Visible;
            
        }

        /// <summary>
        /// Call GetAccessToken - to acquire a token requiring user to sign-in
        /// Then Fetch Email is called using the received token
        /// </summary>
        private async void GetInboxData_Click(object sender, RoutedEventArgs e)
        {
            grdCompose.Visibility = Visibility.Hidden;
            lblLoading.Visibility = Visibility.Visible;

            authResult = await GetAccessToken();

            if (authResult != null)
            {
                lblLoading.Visibility = Visibility.Hidden;
                FetchEmails(authResult.AccessToken);
                grdInbox.Visibility = Visibility.Visible;
            }
        }

        private void FetchEmails(string accessToken)
        {
            var emails = EmailReader.ReadEmails(new Email { Token = accessToken });
            DisplayReceivedEmails(emails);
        }

        /// <summary>
        /// Bind the email data to the Data grid
        /// </summary>
        /// <param name="emails"></param>
        private void DisplayReceivedEmails(List<EmailMessage> emails)
        {
            var list = new List<EmailViewer>();

            foreach (var item in emails)
            {
                list.Add(new EmailViewer
                {
                    On = item.DateTimeReceived,
                    From = item.From.Address,
                    Subject = item.Subject,
                    Message = item.TextBody,
                    To = string.Join(",", item.ToRecipients.Select(a => a.Name))
                });
            }

            dgMails.ItemsSource = list;
        }
         

        /// <summary>
        /// Sign out the current user
        /// </summary>
        private async void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            var accounts = await App.PublicClientApp.GetAccountsAsync();

            if (accounts.Any())
            {
                try
                {
                    await App.PublicClientApp.RemoveAsync(accounts.FirstOrDefault());
                    authResult = null;
                    this.GetInboxData.Visibility = Visibility.Visible;
                    this.SignOutButton.Visibility = Visibility.Collapsed;
                }
                catch (MsalException ex)
                {                   
                    MessageBox.Show(ex.Message, "", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
        }


        private void ComposeEmail_Click(object sender, RoutedEventArgs e)
        {
            grdInbox.Visibility = Visibility.Collapsed;
            grdCompose.Visibility = Visibility.Visible;
        }

        private async void btnSend_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateForm())
                return;

            if (authResult == null)
            {
                authResult = await GetAccessToken();
            }

            Email email = new Email
            {
                ToRecipients = GetAddresses(),
                Subject = txtSubject.Text.Trim(),
                Body = txtContent.Text.Trim(),
                Token = authResult?.AccessToken
            };

            if (email.ToRecipients == null || !email.ToRecipients.Any())
            {
                MessageBox.Show("Invalid Email Id! Please enter a valid one", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            EmailSender.SendEmail(email);

            MessageBox.Show("Email Sent Successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

            ClearComposeForm();

        }

        private bool ValidateForm()
        {
            if (txtTo.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Email Id", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtTo.Focus();
                return false;
            }
            else if (txtSubject.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Subject", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtSubject.Focus();
                return false;
            }
            else if (txtContent.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter your message", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtContent.Focus();
                return false;
            }

            return true;
        }

        private async Task<AuthenticationResult> GetAccessToken()
        {
            var app = App.PublicClientApp;

            var accounts = await app.GetAccountsAsync();
            var firstAccount = accounts.FirstOrDefault();

            try
            {
                authResult = await app.AcquireTokenSilent(scopes, firstAccount)
                    .ExecuteAsync();
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilent. 
                // This indicates you need to call AcquireTokenInteractive to acquire a token
                System.Diagnostics.Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

                try
                {
                    authResult = await app.AcquireTokenInteractive(scopes)
                        .WithAccount(accounts.FirstOrDefault())
                        .WithParentActivityOrWindow(new WindowInteropHelper(this).Handle) // optional, used to center the browser on the window
                        .WithPrompt(Prompt.SelectAccount)
                        .ExecuteAsync();
                }
                catch (MsalException msalex)
                { 
                    MessageBox.Show(msalex.Message, "Error Acquiring Token", MessageBoxButton.OK, MessageBoxImage.Hand);
                }
            }
            catch (Exception ex)
            { 
                MessageBox.Show(ex.Message, "Error Acquiring Token", MessageBoxButton.OK, MessageBoxImage.Hand);
            }

            if(authResult!=null)
                this.SignOutButton.Visibility = Visibility.Visible;

            return authResult;
        }

        private void ClearComposeForm()
        {
            txtContent.Clear();
            txtSubject.Clear();
            txtTo.Clear();
        }


        /// <summary>
        /// Accept and Validate each Email ID entered by user. Invalid's are ignored
        /// </summary>
        /// <returns>List of valid email ids </returns>
        private List<EmailAddress> GetAddresses()
        {
            var addresses = txtTo.Text.Trim();

            var validEmails = new List<EmailAddress>();

            if (!addresses.Contains(","))
            {
                if (!Utils.IsValidEmail(addresses))
                {
                    return null;
                }

                validEmails.Add(new EmailAddress(addresses));

                return validEmails;
            }

            var emails = addresses.Split(',');
            var invalidEmails = "";

            foreach (var emailAddress in emails)
            {
                if (Utils.IsValidEmail(emailAddress.Trim()))
                {
                    validEmails.Add(new EmailAddress(emailAddress));
                }
                else
                    invalidEmails +=  emailAddress + Environment.NewLine;
            }

            if (invalidEmails != "")
                MessageBox.Show("These invalid Ids are ignored " + Environment.NewLine + invalidEmails, "Ignored Email Ids", MessageBoxButton.OK, MessageBoxImage.Information);

            return validEmails;
        }

        private void btnReset_Click(object sender, RoutedEventArgs e)
        {
            ClearComposeForm();
        } 

       
    }
}
