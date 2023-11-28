using System.Collections.Generic;

namespace Frends.Exchange.SendEmail.Definitions;

/// <summary>
/// Result.
/// </summary>
public class Result
{
    /// <summary>
    /// Gets a value indicating whether Task was executed successfully.
    /// </summary>
    /// <example>true</example>
    public bool Success { get; private set; }

    public List<Output> Data { get; private set; }

    internal Result(bool success, List<Output> data)
    {
        Success = success;
        Data = data;
    }
}
