﻿using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Exchange.SendEmail.Definitions;

/// <summary>
/// Attachment parameters.
/// </summary>
public class Attachments
{
    /// <summary>
    /// Chooses if the attachment file is created from a string or copied from disk. 
    /// This determines how the attachment is added to the email.
    /// </summary>
    /// <example>AttachmentTypes.FileAttachment</example>
    [DefaultValue(AttachmentTypes.FileAttachment)]
    public AttachmentTypes AttachmentType { get; set; }

    /// <summary>
    /// Name of the attachment file.
    /// </summary>
    /// <example>message.txt</example>
    [UIHint(nameof(AttachmentType), "", AttachmentTypes.AttachmentFromString)]
    public string FileName { get; set; }

    /// <summary>
    /// Content of the attachment file. 
    /// </summary>
    /// <example>Hello, World!</example>
    [UIHint(nameof(AttachmentType), "", AttachmentTypes.AttachmentFromString)]
    public string FileContent { get; set; }

    /// <summary>
    /// Attachment file's path. 
    /// If the path ends in a directory, all files in that folder with given Attachments.FileMask are added as attachments.
    /// </summary>
    /// <example>C:\temp\message.txt</example>
    [UIHint(nameof(AttachmentType), "", AttachmentTypes.FileAttachment)]
    public string FilePath { get; set; }

    /// <summary>
    /// Filemask. 
    /// </summary>
    /// <example>*.txt</example>
    [UIHint(nameof(AttachmentType), "", AttachmentTypes.FileAttachment)]
    public string FileMask { get; set; }
}
