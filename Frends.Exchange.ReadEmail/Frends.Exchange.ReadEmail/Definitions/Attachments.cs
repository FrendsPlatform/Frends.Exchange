namespace Frends.Exchange.ReadEmail.Definitions;

/// <summary>
/// Represents an attachment.
/// </summary>
public class Attachments
{
    /// <summary>
    /// Specifies the ID of the attachment.
    /// </summary>
    /// <example>AAMkADIxYTJiZDIz...</example>
    public string Id { get; set; }

    /// <summary>
    /// Specifies the name of the attachment.
    /// </summary>
    /// <example>C:\temp\file.txt</example>
    public string FilePath { get; set; }

    /// <summary>
    /// Specifies the size of the attachment in bytes.
    /// </summary>
    /// <example>6000</example>
    public int? Size { get; set; }

    /// <summary>
    /// Specifies the data type of the attachment.
    /// </summary>
    /// <example>#microsoft.graph.fileAttachment</example>
    public string OdataType { get; set; }

    /// <summary>
    /// Specifies the content of the attachment.
    /// </summary>
    /// <example>This is content.</example>
    public string Content { get; set; }
}
