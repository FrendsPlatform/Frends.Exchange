using System.ComponentModel;

namespace Frends.Exchange.SendEmail.Definitions;

/// <summary>
/// Email content.
/// </summary>
public class Input
{
    /// <summary>
    /// Sender email address. 
    /// This is the email address that will appear in the "From" field of the email.
    /// </summary>
    /// <example>sender@example.com</example>
    public string From { get; set; }

    /// <summary>
    /// Recipient addresses separated by ',' or ';'. 
    /// These are the main recipients of the email.
    /// </summary>
    /// <example>recipient1@example.com, recipient2@example.com</example>
    public string To { get; set; }

    /// <summary>
    /// Cc recipient addresses separated by ',' or ';'. 
    /// These recipients will receive a copy of the email.
    /// </summary>
    /// <example>cc1@example.com, cc2@example.com</example>
    public string Cc { get; set; }

    /// <summary>
    /// Bcc recipient addresses separated by ',' or ';'. 
    /// These recipients will receive a copy of the email, but other recipients will not see their email addresses.
    /// </summary>
    /// <example>bcc1@example.com, bcc2@example.com</example>
    public string Bcc { get; set; }

    /// <summary>
    /// Email message's subject. 
    /// This is the main topic or title of the email.
    /// </summary>
    /// <example>Meeting Reminder</example>
    public string Subject { get; set; }

    /// <summary>
    /// Body of the message. 
    /// This is the main content of the email.
    /// </summary>
    /// <example>Don't forget about our meeting tomorrow at 10am.</example>
    public string Message { get; set; }

    /// <summary>
    /// Set this true if the message is HTML. 
    /// This allows you to send emails with HTML formatting.
    /// </summary>
    /// <example>false</example>
    [DefaultValue(false)]
    public bool IsMessageHtml { get; set; }

    /// <summary>
    /// Importance level. 
    /// This sets the importance level of the email (e.g., Low, Normal, High).
    /// </summary>
    /// <example>ImportanceLevels.Normal</example>
    [DefaultValue(ImportanceLevels.Normal)]
    public ImportanceLevels Importance { get; set; }

    /// <summary>
    /// Indicates whether to save the message in Sent Items.
    /// </summary>
    /// <example>true</example>
    [DefaultValue(true)]
    public bool SaveToSentItems { get; set; }

    /// <summary>
    /// Email attachments.
    /// </summary>
    /// <example>
    ///     { AttachmentTypes.FileAttachment, C:\temp\temp.txt, *.* }, 
    ///     { AttachmentTypes.AttachmentFromString, temp.txt, "This is temp file." }
    /// </example>
    public Attachments[] Attachments { get; set; }
}