﻿using Frends.Exchange.ReadEmail.Definitions;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

namespace Frends.Exchange.ReadEmail
{
    /// <summary>
    /// A collection of methods that are public to ensure testing,
    /// but in a separate class to hide them from UI.
    /// </summary>
    public static class Util
    {
        /// <summary>
        /// Save attachments from collection to files.
        /// </summary>
        /// <returns>List of full paths to saved file attachments.</returns>
        public static List<string> SaveAttachments(AttachmentCollection attachments, ExchangeOptions options)
        {
            List<string> pathList = new() { };

            foreach (var attachment in attachments)
            {
                FileAttachment file = attachment as FileAttachment;
                string path = Path.Combine(
                    options.AttachmentSaveDirectory,
                    options.OverwriteAttachment ? file.Name :
                        string.Concat(
                            Path.GetFileNameWithoutExtension(file.Name), "_",
                            Guid.NewGuid().ToString(),
                            Path.GetExtension(file.Name))
                        );
                file.Load(path);
                pathList.Add(path);
            }

            return pathList;
        }

        /// <summary>
        ///     As Per MSDN Example, to ensure SSL. Copy and Paste.
        ///     https://msdn.microsoft.com/en-us/library/office/dd633677(v=exchg.80).aspx
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="certificate"></param>
        /// <param name="chain"></param>
        /// <param name="sslPolicyErrors"></param>
        /// <returns>bool</returns>
        public static bool ExchangeCertificateValidationCallBack(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == SslPolicyErrors.None) return true;

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (var status in chain.ChainStatus)
                    {
                        if (status.Status != X509ChainStatusFlags.NoError)
                        {
                            return false;
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are untrusted root errors for self-signed certificates.
                // These certificates are valid for default Exchange server installations, so return true.
                return true;
            }
            // In all other cases, return false.
            return false;
        }

        /// <summary>
        /// Helper for connecting to Exchange service.
        /// </summary>
        /// <param name="settings">Exchange server related settings</param>
        /// <returns></returns>
        public static ExchangeService ConnectToExchangeService(ExchangeSettings settings)
        {
            ExchangeVersion ev;
            var office365 = false;
            switch (settings.ExchangeServerVersion)
            {
                case ExchangeServerVersion.Exchange2007_SP1:
                    ev = ExchangeVersion.Exchange2007_SP1;
                    break;
                case ExchangeServerVersion.Exchange2010:
                    ev = ExchangeVersion.Exchange2010;
                    break;
                case ExchangeServerVersion.Exchange2010_SP1:
                    ev = ExchangeVersion.Exchange2010_SP1;
                    break;
                case ExchangeServerVersion.Exchange2010_SP2:
                    ev = ExchangeVersion.Exchange2010_SP2;
                    break;
                case ExchangeServerVersion.Exchange2013:
                    ev = ExchangeVersion.Exchange2013;
                    break;
                case ExchangeServerVersion.Exchange2013_SP1:
                    ev = ExchangeVersion.Exchange2013_SP1;
                    break;
                case ExchangeServerVersion.Office365:
                    ev = ExchangeVersion.Exchange2013_SP1;
                    office365 = true;
                    break;
                default:
                    ev = ExchangeVersion.Exchange2013;
                    break;
            }

            var service = new ExchangeService(ev);

            // SSL certification check.
            ServicePointManager.ServerCertificateValidationCallback = ExchangeCertificateValidationCallBack;

            if (!office365)
            {
                if (string.IsNullOrWhiteSpace(settings.Username)) service.UseDefaultCredentials = true;
                else service.Credentials = new NetworkCredential(settings.Username, settings.Password);
            }
            else service.Credentials = new WebCredentials(settings.Username, settings.Password);

            if (settings.UseAutoDiscover) service.AutodiscoverUrl(settings.Username, RedirectionUrlValidationCallback);
            else service.Url = new Uri(settings.ServerAddress);

            return service;
        }

        // The following is a basic redirection validation callback method.
        // It inspects the redirection URL and only allows the Service object to follow the redirection link if the URL is using HTTPS. 
        // This redirection URL validation callback provides sufficient security for development and testing of your application.
        // However, it may not provide sufficient security for your deployed application.
        // You should always make sure that the URL validation callback method that you use meets the security requirements of your organization.
        /// <summary>
        /// Returns true if the provided url has https scheme. Otherwise returns false.
        /// </summary>
        public static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            var result = false;
            var redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL.
            // In this simple validation callback, the redirection URL is considered valid if it is using HTTPS to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https") result = true;

            return result;
        }

    }
}
