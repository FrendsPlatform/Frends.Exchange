using System.ComponentModel;

namespace Frends.Exchange.SendEmail.Definitions;

/// <summary>
/// Options parameters.
/// </summary>
public class Options
{
    /// <summary>
    /// Value indicating whether an error should stop the Task and throw an exception (true) or try to continue processing and add a new object to Result.Data list as { EmailSent = false, StatusString = "Error message." }
    /// </summary>
    /// <example>true</example>
    [DefaultValue(true)]
    public bool ThrowExceptionOnFailure { get; set; }

    /// <summary>
    /// If set to true and no files match the given path, an exception is thrown. 
    /// This is used to ensure that the email has the correct attachments before being sent.
    /// </summary>
    /// <example>true</example>
    public bool ThrowExceptionIfAttachmentNotFound { get; set; }
}