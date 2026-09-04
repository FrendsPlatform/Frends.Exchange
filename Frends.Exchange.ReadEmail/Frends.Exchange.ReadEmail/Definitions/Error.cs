namespace Frends.Exchange.ReadEmail.Definitions;

/// <summary>
/// Represents an error that occurred during the task execution.
/// </summary>
public class Error
{
    /// <summary>
    /// Error message.
    /// </summary>
    /// <example>One or more required connection values missing: TenantId, ClientId, ClientSecret.</example>
    public string Message { get; set; }

    /// <summary>
    /// Additional information about the error, such as the original exception.
    /// </summary>
    /// <example>System.ArgumentNullException: Value cannot be null.</example>
    public dynamic AdditionalInfo { get; set; }
}
