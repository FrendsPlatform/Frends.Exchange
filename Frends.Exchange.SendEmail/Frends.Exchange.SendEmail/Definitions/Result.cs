namespace Frends.Exchange.SendEmail.Definitions;

/// <summary>
/// Represents the result of a task.
/// </summary>
public class Result
{
    /// <summary>
    /// Gets a value indicating whether the task was executed successfully.
    /// </summary>
    /// <example>true</example>
    public bool Success { get; private set; }

    /// <summary>
    /// Gets the result of the task. Contains exception message if exception was thrown and Options.ThrowExceptionOnFailure = false.
    /// </summary>
    /// <example>Email sent successfully.</example>
    public string Data { get; private set; }

    internal Result(bool success, string data)
    {
        Success = success;
        Data = data;
    }
}
