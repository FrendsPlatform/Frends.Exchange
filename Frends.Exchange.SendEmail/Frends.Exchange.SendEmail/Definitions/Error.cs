namespace Frends.Exchange.SendEmail.Definitions;

/// <summary>
/// Represents error details of a failed task execution.
/// </summary>
public class Error
{
    /// <summary>
    /// Gets or sets the error message.
    /// </summary>
    /// <example>Failed to send an email. Object reference not set to an instance of an object.</example>
    public string Message { get; set; }

    /// <summary>
    /// Gets or sets additional information about the error, such as the original exception.
    /// </summary>
    /// <example>System.Exception: Object reference not set to an instance of an object.</example>
    public object AdditionalInfo { get; set; }
}
