using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Exchange.ReadEmail.Definitions;

/// <summary>
/// Email content.
/// </summary>
public class Input
{
    /// <summary>
    /// User entity to get emails from. Can be left empty to use default user. 
    /// </summary>
    /// <example>johndoe@example.com</example>
    public string From { get; set; }

    /// <summary>
    /// Select properties to be returned.
    /// </summary>
    /// <example>subject,body,bodyPreview,uniqueBody,hasAttachments</example>
    public string Select { get; set; }

    /// <summary>
    /// Filter items by property values.
    /// </summary>
    /// <example>parentFolderId eq 'INBOX' and from/emailAddress/address eq 'user@exampledomain.com' and subject eq 'This is subject' and isRead eq true and hasAttachments eq true</example>
    public string Filter { get; set; }

    /// <summary>
    /// Skip the first n items.
    /// </summary>
    /// <example>10</example>
    [DefaultValue(null)]
    public int? Skip { get; set; }

    /// <summary>
    /// Limits the result to the first n items. If set to 0, all items are returned.
    /// </summary>
    /// <example>10</example>
    [DefaultValue(0)]
    public int Top { get; set; }

    /// <summary>
    /// Order items by property values.
    /// </summary>
    /// <example>receivedDateTime DESC,subject ASC</example>
    public string Orderby { get; set; }

    /// <summary>
    /// Expand related entities.
    /// </summary>
    /// <example>attachments</example>
    public string Expand { get; set; }

    /// <summary>
    /// Specifies whether to mark an email as read after processing it. If set to true, the email will be marked as read in the mailbox. If set to false, the email's read status will remain unchanged.
    /// </summary>
    /// <example>true</example>
    [DefaultValue(true)]
    public bool UpdateReadStatus { get; set; }

    /// <summary>
    /// Header parameters.
    /// </summary>
    /// <example>{ [ "Prefer", "outlook.body-content-type=\"text\"" ] }</example>
    public HeaderParameters[] Headers { get; set; }

    /// <summary>
    /// Specifies whether to download attachments.
    /// </summary>
    /// <example>true</example>
    [DefaultValue(true)]
    public bool DownloadAttachments { get; set; }

    /// <summary>
    /// Specifies the directory where the downloaded attachments will be stored.
    /// </summary>
    /// <example>C:\\Users\\username\\Downloads</example>
    [UIHint(nameof(DownloadAttachments), "", true)]
    public string DestinationDirectory { get; set; }

    /// <summary>
    /// Create directory where the downloaded attachments will be stored.
    /// </summary>
    /// <example></example>
    [UIHint(nameof(DownloadAttachments), "", true)]
    public bool CreateDirectory { get; set; }

    /// <summary>
    /// Specifies the action to take when a file with the same name already exists in the destination directory.
    /// If FileExistHandler.Rename is selected, an unique number will be appended to the name of the file to be downloaded (e.g., file(2).txt).
    /// </summary>
    /// <example>FileExistHandlers.Skip</example>
    [UIHint(nameof(DownloadAttachments), "", true)]
    [DefaultValue(FileExistHandlers.Skip)]
    public FileExistHandlers FileExistHandler { get; set; }
}