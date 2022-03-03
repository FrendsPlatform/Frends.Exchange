using Frends.Exchange.ReadEmail.Definitions;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace Frends.Exchange.ReadEmail;

/// <summary>
/// Tasks for interacting with Microsoft Exchange.
/// </summary>
public static class Exchange
{
    /// <summary>
    /// Reads emails via Microsoft Exchange.
    /// </summary>
    public async static Task<List<EmailMessageResult>> ReadEmail([PropertyTab] ExchangeSettings settings, [PropertyTab] ExchangeOptions options)
    {
        if (!options.IgnoreAttachments)
        {
            if (string.IsNullOrEmpty(options.AttachmentSaveDirectory)) {
                throw new ArgumentNullException("No save directory given.", nameof(ExchangeOptions.AttachmentSaveDirectory));
            }
            if (!Directory.Exists(options.AttachmentSaveDirectory))
            {
                throw new DirectoryNotFoundException($"Could not find or access attachment save directory {options.AttachmentSaveDirectory}");
            }
        }
        if (options.MaxEmails <= 0)
        {
            throw new ArgumentException("MaxEmails can't be lower than 1.");
        }


        // Connect, create view and search filter
        ExchangeService exchangeService = ConnectToExchangeService(settings);
        ItemView view = new(options.MaxEmails);
        var searchFilter = BuildFilterCollection(options);
        FindItemsResults<Item> exchangeResults;

        if (!string.IsNullOrEmpty(settings.Mailbox))
        {
            var mb = new Mailbox(settings.Mailbox);
            var fid = new FolderId(WellKnownFolderName.Inbox, mb);
            var inbox = await Folder.Bind(exchangeService, fid);
            exchangeResults = searchFilter.Count == 0 ? await inbox.FindItems(view) : await inbox.FindItems(searchFilter, view);
        }
        else
        {
            exchangeResults = searchFilter.Count == 0 ? await exchangeService.FindItems(WellKnownFolderName.Inbox, view) : await exchangeService.FindItems(WellKnownFolderName.Inbox, searchFilter, view);
        }
        // Get email items
        var emails = exchangeResults.Where(msg => msg is EmailMessage).Cast<EmailMessage>().ToList();

        // Check if list is empty and if an error needs to be thrown.
        if (emails.Any() && options.ThrowErrorIfNoMessagesFound)
        {
            // If not, return a result with a notification of no found messages.
            throw new ArgumentException("No messages found matching the search filter.",
                paramName: nameof(options.ThrowErrorIfNoMessagesFound));
        }

        // Load properties for each email and process attachments
        var result = ReadEmails(emails, exchangeService, options);

        // should delete mails?
        if (options.DeleteReadEmails)
            emails.ForEach(msg => msg.Delete(DeleteMode.HardDelete));

        // should mark mails as read?
        if (!options.DeleteReadEmails && options.MarkEmailsAsRead)
        {
            foreach (EmailMessage msg in emails)
            {
                msg.IsRead = true;
                await msg.Update(ConflictResolutionMode.AutoResolve);
            }
        }

        return await result;
    }

    /// <summary>
    /// Build search filter from options.
    /// </summary>
    /// <param name="options">Options.</param>
    /// <returns>Search filter collection.</returns>
    internal static SearchFilter.SearchFilterCollection BuildFilterCollection(ExchangeOptions options)
    {
        // Create search filter collection.
        var searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And);

        // Construct rest of search filter based on options
        if (options.GetOnlyEmailsWithAttachments)
            searchFilter.Add(new SearchFilter.IsEqualTo(ItemSchema.HasAttachments, true));

        if (options.GetOnlyUnreadEmails)
            searchFilter.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));

        if (!string.IsNullOrEmpty(options.EmailSenderFilter))
            searchFilter.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.Sender, options.EmailSenderFilter));

        if (!string.IsNullOrEmpty(options.EmailSubjectFilter))
            searchFilter.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, options.EmailSubjectFilter));

        return searchFilter;
    }

    /// <summary>
    /// Convert Email collection t EMailMessageResults.
    /// </summary>
    /// <param name="emails">Emails collection.</param>
    /// <param name="exchangeService">Exchange services.</param>
    /// <param name="options">Options.</param>
    /// <returns>Collection of EmailMessageResult.</returns>
    internal async static Task<List<EmailMessageResult>> ReadEmails(IEnumerable<EmailMessage> emails, ExchangeService exchangeService, ExchangeOptions options)
    {
        List<EmailMessageResult> result = new();

        foreach (EmailMessage email in emails)
        {
            // Define property set
            var propSet = new PropertySet(
                    BasePropertySet.FirstClassProperties,
                    EmailMessageSchema.Body,
                    EmailMessageSchema.Attachments);

            // Bind and load email message with desired properties
            var newEmail = await EmailMessage.Bind(exchangeService, email.Id, propSet);

            var pathList = new List<string>();
            if (!options.IgnoreAttachments)
            {
                // Save all attachments to given directory

                pathList = SaveAttachments(newEmail.Attachments, options);
            }

            // Build result for email message

            var emailMessage = new EmailMessageResult
            {
                Id = newEmail.Id.UniqueId,
                Date = newEmail.DateTimeReceived,
                Subject = newEmail.Subject,
                BodyText = "",
                BodyHtml = newEmail.Body.Text,
                To = string.Join(",", newEmail.ToRecipients.Select(j => j.Address)),
                From = newEmail.From.Address,
                Cc = string.Join(",", newEmail.CcRecipients.Select(j => j.Address)),
                AttachmentSaveDirs = pathList
            };

            // Catch exception in case of server version is earlier than Exchange2013
            try { emailMessage.BodyText = newEmail.TextBody.Text; } catch { }

            result.Add(emailMessage);
        }

        return result;
    }

    /// <summary>
    /// Save attachments from collection to files.
    /// </summary>
    /// <returns>List of full paths to saved file attachments.</returns>
    internal static List<string> SaveAttachments(AttachmentCollection attachments, ExchangeOptions options)
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
    internal static bool ExchangeCertificateValidationCallBack(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
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
    internal static ExchangeService ConnectToExchangeService(ExchangeSettings settings)
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
    internal static bool RedirectionUrlValidationCallback(string redirectionUrl)
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