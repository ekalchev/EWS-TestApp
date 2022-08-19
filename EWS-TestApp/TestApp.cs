using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace EWS_TestApp
{
    internal class TestApp
    {
        private const string username = "emil.kalchev@mobiexch.net";
        private const string password = "Password1";

        public async Task RunAsync()
        {
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;

            //ExchangeService service = CreateServiceUsingAutoDiscovery();
            ExchangeService service = CreateServiceForUrl("https://vm-exchange01.mobiexch.net/EWS/Exchange.asmx");
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

        private ExchangeService CreateServiceForUrl(string url)
        {
            // Create the binding.
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
            // Set the credentials for the on-premises server.
            service.Credentials = new NetworkCredential(username, password);
            // Set the URL.
            service.Url = new Uri(url);
            return service;
        }

        private ExchangeService CreateServiceUsingAutoDiscovery()
        {
            // Create the binding.
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
            // Set the credentials for the on-premises server.
            service.Credentials = new NetworkCredential(username, password);
            // Set the URL.
            service.AutodiscoverUrl(username);
            return service;
        }
    }
}
