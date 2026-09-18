using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Exchange.ReadEmail.Definitions;

/// <summary>
/// Options for controlling the behavior of a task.
/// </summary>
public class Options
{
    /// <summary>
    /// If true, then received emails will be hard deleted.
    /// </summary>
    /// <example>false</example>
    [DefaultValue(false)]
    public bool DeleteReadEmails { get; set; } = false;

    /// <summary>
    /// Gets or sets a value indicating whether an error should throw an exception or return a failed Result.
    /// If set to true, an exception will be thrown when an error occurs. If set to false, execution will stop immediately and return a Result with Success set to false and error details.
    /// </summary>
    /// <example>true</example>
    [DefaultValue(true)]
    public bool ThrowErrorOnFailure { get; set; } = true;

    /// <summary>
    /// Overrides the error message on failure. If `ThrowErrorOnFailure` is set to `true`, then the original exception will be wrapped in a new Exception with this error message.
    /// </summary>
    /// <example>Reading emails from mailbox failed: connection could not be established</example>
    [DisplayFormat(DataFormatString = "Text")]
    [DefaultValue("")]
    public string ErrorMessageOnFailure { get; set; } = string.Empty;
}