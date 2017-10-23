using System;
using System.Threading.Tasks;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace GraphConsoleAppV3
{
    internal class AuthenticationHelper
    {
        public static string TokenForUser;

        /// <summary>
        /// Async task to acquire token for Application.
        /// </summary>
        /// <returns>Async Token for application.</returns>
        public static async Task<string> AcquireTokenAsyncForApplication()
        {
            return  GetTokenForApplication();
        }

        /// <summary>
        /// Get Token for Application.
        /// </summary>
        /// <returns>Token for application.</returns>
        public static string GetTokenForApplication()
        {
            AuthenticationContext authenticationContext = new AuthenticationContext(Constants.AuthString, false);
            // Config for OAuth client credentials 
            //ClientCredential clientCred = new ClientCredential(Constants.ClientId, Constants.ClientSecret);


            //AuthenticationResult authenticationResult = authenticationContext.AcquireToken(Constants.ResourceUrl,
            //    clientCred);
            AuthenticationResult result = null;
            var username = "x1@xuhuadd.partner.onmschina.cn";
            var password = "QWERasd4317";
            // var userPasswordCredential = new UserPasswordCredential(username, password);
            UserCredential uc = new UserCredential(username, password);
            try
            {
                 result = authenticationContext.AcquireToken(Constants.ResourceUrl, Constants.ClientId, uc);
            }
            catch (Exception e) {
                string a = e.Message.ToString();
            }

         //   authenticationContext.AcquireToken()
            string token = result.AccessToken;
            return token;
        }

        /// <summary>
        /// Get Active Directory ClienSt for Application.
        /// </summary>
        /// <returns>ActiveDirectoryClient for Application.</returns>
        public static ActiveDirectoryClient GetActiveDirectoryClientAsApplication()
        {
            Uri servicePointUri = new Uri(Constants.ResourceUrl);
            Uri serviceRoot = new Uri(servicePointUri, Constants.TenantId);
            ActiveDirectoryClient activeDirectoryClient = new ActiveDirectoryClient(serviceRoot,
                async () => await AcquireTokenAsyncForApplication());
            return activeDirectoryClient;
        }

        /// <summary>
        /// Async task to acquire token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> AcquireTokenAsyncForUser()
        {
            return GetTokenForUser();
        }

        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static string GetTokenForUser()
        {
            if (TokenForUser == null)
            {
                var redirectUri = new Uri("https://localhost");
                AuthenticationContext authenticationContext = new AuthenticationContext(Constants.AuthString, false);
                AuthenticationResult userAuthnResult = authenticationContext.AcquireToken(Constants.ResourceUrl,
                    Constants.ClientIdForUserAuthn, redirectUri, PromptBehavior.Always);
               
              
                TokenForUser = userAuthnResult.AccessToken;
                Console.WriteLine("\n Welcome " + userAuthnResult.UserInfo.GivenName + " " +
                                  userAuthnResult.UserInfo.FamilyName);
            }
            return TokenForUser;
        }

        /// <summary>
        /// Get Active Directory Client for User.
        /// </summary>
        /// <returns>ActiveDirectoryClient for User.</returns>
        //public static ActiveDirectoryClient GetActiveDirectoryClientAsUser()
        //{
        //    Uri servicePointUri = new Uri(Constants.ResourceUrl);
        //    Uri serviceRoot = new Uri(servicePointUri, Constants.TenantId);
        //    ActiveDirectoryClient activeDirectoryClient = new ActiveDirectoryClient(serviceRoot,
        //        async () => await AcquireTokenAsyncForUser());
        //    return activeDirectoryClient;
        //}
    }
}
