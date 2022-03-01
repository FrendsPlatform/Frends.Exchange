﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Frends.Exchange.ReadEmail.Definitions
{

    /// <summary>
    /// Exchange server specific options.
    /// </summary>
    public class ExchangeSettings
    {
        /// <summary>
        /// Which exchange server to target?
        /// </summary>
        public ExchangeServerVersion ExchangeServerVersion { get; set; }

        /// <summary>
        /// If true, will try to auto discover server address from user name.
        /// In this cae Host and Port values are not used.
        /// </summary>
        public bool UseAutoDiscover { get; set; }

        /// <summary>
        /// Exchange server address.
        /// </summary>
        [DefaultValue("exchange.frends.com")]
        [DisplayFormat(DataFormatString = "Text")]
        [UIHint(nameof(UseAutoDiscover), "", false)]
        public string ServerAddress { get; set; }

        /// <summary>
        /// Try to login with agent account?
        /// </summary>
        [DefaultValue(false)]
        [Description("Authorize with agent account")]
        public bool UseAgentAccount { get; set; }

        /// <summary>
        /// Email account to use.
        /// </summary>
        [DefaultValue("agent@frends.com")]
        [DisplayFormat(DataFormatString = "Text")]
        public string Username { get; set; }

        /// <summary>
        /// Account password.
        /// </summary>
        [PasswordPropertyText]
        [UIHint(nameof(UseAgentAccount), "", false)]
        public string Password { get; set; }


        /// <summary>
        /// Inbox to read emails from.
        /// If empty reads from default mailbox.
        /// </summary>
        [DefaultValue("agentinbox@frends.com")]
        [DisplayFormat(DataFormatString = "Text")]
        public string Mailbox { get; set; }
    }


    /// <summary>
    /// Options related to Exchange reading.
    /// </summary>
    public class ExchangeOptions
    {
        /// <summary>
        /// Maximum number of emails to retrieve.
        /// </summary>
        [DefaultValue(10)]
        public int MaxEmails { get; set; }

        /// <summary>
        /// Should get only unread emails?
        /// </summary>
        public bool GetOnlyUnreadEmails { get; set; }

        /// <summary>
        /// If true, then marks queried emails as read.
        /// </summary>
        public bool MarkEmailsAsRead { get; set; }

        /// <summary>
        /// If true, then received emails will be hard deleted.
        /// </summary>
        public bool DeleteReadEmails { get; set; }

        /// <summary>
        /// Optional.
        /// If a sender is given, it will be used to filter emails.
        /// </summary>
        [DefaultValue("")]
        [DisplayFormat(DataFormatString = "Text")]
        public string EmailSenderFilter { get; set; }

        /// <summary>
        /// Optional.
        /// If a subject is given, it will be used to filter emails.
        /// </summary>
        [DefaultValue("")]
        [DisplayFormat(DataFormatString = "Text")]
        public string EmailSubjectFilter { get; set; }

        /// <summary>
        /// If true, the task throws an error if no messages matching search criteria were found.
        /// </summary>
        public bool ThrowErrorIfNoMessagesFound { get; set; }

        /// <summary>
        /// If true, the task doesn't handle emails attachments.
        /// </summary>
        public bool IgnoreAttachments { get; set; }

        /// <summary>
        /// If true, the task fetches only emails with attachments.
        /// </summary>
        [UIHint(nameof(IgnoreAttachments), "", false)]
        public bool GetOnlyEmailsWithAttachments { get; set; }

        /// <summary>
        /// Directory where attachments will be saved to.
        /// </summary>
        [DefaultValue("")]
        [DisplayFormat(DataFormatString = "Text")]
        [UIHint(nameof(IgnoreAttachments), "", false)]
        public string AttachmentSaveDirectory { get; set; }

        /// <summary>
        /// Should the attachment be overwritten, if the save directory already contains an attachment with the same name?
        /// If no, a GUID will be added to the filename.
        /// </summary>
        [UIHint(nameof(IgnoreAttachments), "", false)]
        public bool OverwriteAttachment { get; set; }
    }


    /// <summary>
    /// Output result for read operation.
    /// </summary>
    public class EmailMessageResult
    {
        /// <summary>
        /// Identifier for the email.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Recipient of the email.
        /// </summary>
        public string To { get; set; }

        /// <summary>
        /// Carbon Copy-field of the email.
        /// </summary>
        public string Cc { get; set; }

        /// <summary>
        /// Sender of the email.
        /// </summary>
        public string From { get; set; }

        /// <summary>
        /// Email received date.
        /// </summary>
        public DateTime Date { get; set; }

        /// <summary>
        /// Title of the email.
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Body of the email as text.
        /// </summary>
        public string BodyText { get; set; }

        /// <summary>
        /// Body HTML, if available.
        /// </summary>
        public string BodyHtml { get; set; }

        /// <summary>
        /// Attachment download path.
        /// </summary>
        public List<string> AttachmentSaveDirs { get; set; }
    }

    /// <summary>
    /// Depicts a certain Exchange Server version.
    /// </summary>
    public enum ExchangeServerVersion
    {
#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
        Exchange2007_SP1,
        Exchange2010,
        Exchange2010_SP1,
        Exchange2010_SP2,
        Exchange2013,
        Exchange2013_SP1,
        Office365
#pragma warning restore CS1591 // Missing XML comment for publicly visible type or member
    }
}
