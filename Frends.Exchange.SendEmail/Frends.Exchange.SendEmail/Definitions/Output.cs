namespace Frends.Exchange.SendEmail.Definitions;

/// <summary>
/// Result of an email sending operation.
/// </summary>
public class Output
{
    /// <summary>
    /// Value is true if the email was sent successfully.
    /// </summary>
    /// <example>true</example>
    public bool EmailSent { get; set; }

    /// <summary>
    /// Contains information about the operation.
    /// </summary>
    /// <example>"Message was sent"</example>
    public string MessageStatus { get; set; }
}