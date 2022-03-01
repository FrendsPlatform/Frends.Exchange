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
        ExchangeService exchangeService = Util.ConnectToExchangeService(settings);
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
    private static SearchFilter.SearchFilterCollection BuildFilterCollection(ExchangeOptions options)
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
    private async static Task<List<EmailMessageResult>> ReadEmails(IEnumerable<EmailMessage> emails, ExchangeService exchangeService, ExchangeOptions options)
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

                pathList = Util.SaveAttachments(newEmail.Attachments, options);
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

}