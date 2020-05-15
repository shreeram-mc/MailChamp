using Microsoft.Identity.Client;
using System;
using System.Windows;

namespace active_directory_wpf_msgraph_v2
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    
    // To change from Microsoft public cloud to a national cloud, use another value of AzureCloudInstance
    public partial class App : Application
    {
        static App()
        {
            var ClientId = Environment.GetEnvironmentVariable("AZ_ClientID");
            var Tenant = Environment.GetEnvironmentVariable("AZ_TenantId");

            _clientApp = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority($"{Instance}{Tenant}")
                .WithDefaultRedirectUri()
                .Build();

            TokenCacheHelper.EnableSerialization(_clientApp.UserTokenCache);
        } 
       
        private static string Instance = "https://login.microsoftonline.com/";
        private static IPublicClientApplication _clientApp ;

        public static IPublicClientApplication PublicClientApp { get { return _clientApp; } }
    }
}
