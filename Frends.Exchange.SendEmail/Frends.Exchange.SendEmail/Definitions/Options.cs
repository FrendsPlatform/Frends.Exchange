using System.ComponentModel;

namespace Frends.Exchange.SendEmail.Definitions;

/// <summary>
/// Options for controlling the behavior of a task.
/// </summary>
public class Options
{
    /// <summary>
    /// Gets or sets a value indicating whether an error should stop the task and throw an exception.
    /// </summary>
    /// <remarks>
    /// If set to true, an exception will be thrown when an error occurs. If set to false, the error message will be inserted into Result.Data and Result.Success will be set to false.
    /// </remarks>
    /// <example>true</example>
    [DefaultValue(true)]
    public bool ThrowExceptionOnFailure { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether an exception should be thrown if no files match the given path.
    /// </summary>
    /// <remarks>
    /// This option is used to ensure that the email has the correct attachments before being sent. If set to true and no files match the given path, an exception will be thrown. If set to false, the task will continue without attachments.
    /// </remarks>
    /// <example>true</example>
    public bool ThrowExceptionIfAttachmentNotFound { get; set; }
}