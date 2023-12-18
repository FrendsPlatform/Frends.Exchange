using System.Collections.Generic;

namespace Frends.Exchange.ReadEmail.Definitions;

/// <summary>
/// Represents the result of an operation.
/// </summary>
public class ResultObject
{
    /// <summary>
    /// The unique identifier of the result object.
    /// </summary>
    /// <example>"AAMkADIxY..."</example>
    public string Id { get; set; }

    /// <summary>
    /// The unique identifier of the parent folder.
    /// </summary>
    /// <example>AAMkADIxYTJ...</example>
    public string ParentFolderId { get; set; }

    /// <summary>
    /// From email address.
    /// </summary>
    /// <example>johndoe@example.com</example>
    public string From { get; set; }

    /// <summary>
    /// The email address of the sender.
    /// </summary>
    /// <example>johndoe@example.com</example>
    public string Sender { get; set; }

    /// <summary>
    /// The list of email addresses of the recipients.
    /// </summary>
    /// <example>{ "bar@exampledomain.com", "foo@exampledomain.com" }</example>
    public List<string> ToRecipients { get; set; }

    /// <summary>
    /// The list of email addresses of the CC recipients.
    /// </summary>
    /// <example>{ "bar@exampledomain.com", "foo@exampledomain.com" }</example>
    public List<string> CcRecipients { get; set; }

    /// <summary>
    /// The list of email addresses of the BCC recipients.
    /// </summary>
    /// <example>{ "bar@exampledomain.com", "foo@exampledomain.com" }</example>
    public List<string> BccRecipients { get; set; }

    /// <summary>
    /// The list of email addresses to reply to.
    /// </summary>
    /// <example>{ "bar@exampledomain.com", "foo@exampledomain.com" }</example>
    public List<string> ReplyTo { get; set; }

    /// <summary>
    /// The subject of the message.
    /// </summary>
    /// <example>This is subject.</example>
    public string Subject { get; set; }

    /// <summary>
    /// The content type of the message.
    /// </summary>
    /// <example>HTML</example>
    public string ContentType { get; set; }

    /// <summary>
    /// The content of the message.
    /// </summary>
    /// <example>This is message's content.</example>
    public string Content { get; set; }

    /// <summary>
    /// The categories of the message.
    /// </summary>
    /// <example>{ "Blue Category" }</example>
    public List<string> Categories { get; set; }

    /// <summary>
    /// The importance of the message.
    /// </summary>
    /// <example>Normal</example>
    public string Importance { get; set; }

    /// <summary>
    /// Indicates whether the message is a draft.
    /// </summary>
    /// <example>false</example>
    public bool IsDraft { get; set; }

    /// <summary>
    /// Indicates whether the message has been read.
    /// </summary>
    /// <example>false</example>
    public bool IsRead { get; set; }

    /// <summary>
    /// Indicates whether the message has attachments.
    /// </summary>
    /// <example>true</example>
    public bool HasAttachments { get; set; }

    /// <summary>
    /// The list of extensions of the message.
    /// </summary>
    /// <example>{ "foo", "bar" }</example>
    public List<string> Extensions { get; set; }

    /// <summary>
    /// The list of attachments of the message.
    /// </summary>
    /// <example>
    /// { 
    ///     { "AAMkADIxYTJiZDIz...", "C:\temp\file.txt", 6000, "#microsoft.graph.itemAttachment", "This is content" },
    ///     { "RjZTNiNwBGAAAAaa...", "C:\temp\file2.txt", 6001, "#microsoft.graph.fileAttachment", "#microsoft.graph.itemAttachment" }
    /// }
    /// </example>
    public List<Attachments> Attachments { get; set; }
}
