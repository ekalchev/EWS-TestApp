using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace EWS_TestApp
{
    // Autodiscover overview: https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/autodiscover-for-exchange

    // OAUTH:
    // OAuth authentication examples: https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth
    // Configure app for OAuth: https://www.michev.info/Blog/Post/3180/exchange-api-permissions-missing
    // To be able to use OAuth with user and password, set allowPublicClient to true in the app manifest

    internal class TestApp
    {
        class Settings
        {
            public string ExchangeUrl { get; private set; }
            public string AutoDiscoverUrl { get; private set; }
            public string UserName { get; private set; }
            public string Password { get; private set; }

            private Settings()
            {
            }

            public static Settings Outlook =>
                new Settings()
                {
                    ExchangeUrl = "https://outlook.office365.com/EWS/Exchange.asmx",
                    AutoDiscoverUrl = "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc",
                    UserName = "neven.mollov@mobisystems.com",
                    Password = "Password1",
                };

            public static Settings MobiTest =>
                new Settings()
                {
                    ExchangeUrl = "https://vm-exchange01.mobiexch.net/EWS/Exchange.asmx",
                    AutoDiscoverUrl = "https://vm-exchange01.mobiexch.net/autodiscover/autodiscover.svc",
                    UserName = "neven.mollov@mobiexch.net",
                    Password = "Password1",
                };
        }

        private static readonly Settings settings = Settings.Outlook;
        private static readonly string clientId = "fffb71af-18e3-4e7b-886d-63eb471f8ee5";
        private static readonly string clientSecret = "UdT8Q~sNgZJ810373wZ7w~uANtKRwTiFDKU5laTL";
        private static readonly string tenantId = "fe426553-582a-4aae-916d-11edb985b6c1";

        public async Task RunAsync()
        {
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;

            ExchangeService service = CreateServiceUsingAutoDiscovery();
            //ExchangeService service = CreateServiceForUrl();
            //ExchangeService service = await CreateServiceUsingOAuthToken();
            //ExchangeService service = await CreateServiceUsingOAuthToken2();

            var properties = new PropertySet(BasePropertySet.IdOnly,
                                                FolderSchema.DisplayName,
                                                new ExtendedPropertyDefinition(0x10f4, MapiPropertyType.Boolean));
            service.BeginSubscribeToStreamingNotificationsOnAllFolders(a =>
            {
                var subscription = service.EndSubscribeToStreamingNotifications(a);
                var connection = new StreamingSubscriptionConnection(service, 30);
                connection.AddSubscription(subscription);

                // Delegate handlers
                connection.OnNotificationEvent += Connection_OnNotificationEvent;
                connection.OnSubscriptionError += Connection_OnSubscriptionError;
                connection.OnDisconnect += Connection_OnDisconnect;
                connection.Open();

            }, null, new EventType[] { EventType.NewMail });

            service.BeginSyncFolderHierarchy(ar =>
            {
                try
                {
                    var result = service.EndSyncFolderHierarchy(ar);
                    //tcs.TrySetResult(result.Where(x => x.ChangeType == ChangeType.Create).Select(x => x.Folder).ToList());
                }
                catch (OperationCanceledException)
                {
                    //tcs.TrySetCanceled();
                }
                catch (Exception ex)
                {
                    //tcs.TrySetException(ex);
                }
            }, null, new FolderId(WellKnownFolderName.Root), properties, null);


            await Task.Delay(100000);
        }

        private void Connection_OnDisconnect(object sender, SubscriptionErrorEventArgs args)
        {

        }

        private void Connection_OnSubscriptionError(object sender, SubscriptionErrorEventArgs args)
        {

        }

        private void Connection_OnNotificationEvent(object sender, NotificationEventArgs args)
        {
        }

        private static bool CertificateValidationCallBack(
         object sender,
         System.Security.Cryptography.X509Certificates.X509Certificate certificate,
         System.Security.Cryptography.X509Certificates.X509Chain chain,
         System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }
            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                        (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }
                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }

        private ExchangeService CreateServiceForUrl()
        {
            // Create the binding.
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
            // Set the credentials for the on-premises server.
            service.Credentials = new NetworkCredential(settings.UserName, settings.Password);
            // Set the URL.
            service.Url = new Uri(settings.ExchangeUrl);
            return service;
        }

        private ExchangeService CreateServiceUsingAutoDiscovery()
        {
            try
            {
                // Create the binding.
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
                // Set the credentials for the on-premises server.
                service.Credentials = new NetworkCredential(settings.UserName, settings.Password);
                // Set the URL.
                service.AutodiscoverUrl(settings.UserName);
                return service;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// App-only authentication
        /// </summary>
        /// <returns></returns>
        private async Task<ExchangeService> CreateServiceUsingOAuthToken()
        {
            try
            {
                // Using Microsoft.Identity.Client 4.22.0
                var cca = ConfidentialClientApplicationBuilder
                    .Create(clientId)
                    .WithClientSecret(clientSecret)
                    .WithTenantId(tenantId)
                    .Build();

                // The permission scope required for EWS access
                var ewsScopes = new string[] { "https://outlook.office365.com/.default" };

                //Make the token request
                var authResult = await cca.AcquireTokenForClient(ewsScopes).ExecuteAsync();



                // Configure the ExchangeService with the access token
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
                service.Url = new Uri(settings.ExchangeUrl);
                service.Credentials = new OAuthCredentials(authResult.AccessToken);

                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, settings.UserName);
                //Include x-anchormailbox header
                service.HttpHeaders.Add("X-AnchorMailbox", settings.UserName);

                var folders = service.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(10));

                return service;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Delegated authentication
        /// </summary>
        /// <returns></returns>
        private async Task<ExchangeService> CreateServiceUsingOAuthToken2()
        {
            try
            {
                // Using Microsoft.Identity.Client 4.22.0

                // Configure the MSAL client to get tokens
                var pcaOptions = new PublicClientApplicationOptions
                {
                    ClientId = clientId,
                    TenantId = tenantId,
                    RedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient",
                };

                var pca = PublicClientApplicationBuilder
                    .CreateWithApplicationOptions(pcaOptions).Build();

                // The permission scope required for EWS access
                var ewsScopes = new string[] { "https://outlook.office365.com/EWS.AccessAsUser.All" };

                // Make the interactive token request
                var authResult = await pca.AcquireTokenInteractive(ewsScopes).ExecuteAsync();



                // Configure the ExchangeService with the access token
                var service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
                service.Url = new Uri(settings.ExchangeUrl);
                service.Credentials = new OAuthCredentials(authResult.AccessToken);

                var folders = service.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(10));

                return service;
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
