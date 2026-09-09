using System.Collections.Generic;

namespace Frends.Exchange.ReadEmail.Definitions;

/// <summary>
/// Represents the result of a task.
/// </summary>
public class Result
{
    internal Result(bool success, List<ResultObject> data, Error error = null)
    {
        Success = success;
        Data = data;
        Error = error;
    }

    /// <summary>
    /// Gets a value indicating whether the task was executed successfully.
    /// </summary>
    /// <example>true</example>
    public bool Success { get; private set; }

    /// <summary>
    /// Gets the result of the task. Contains exception message if exception was thrown and Options.ThrowErrorOnFailure = false.
    /// </summary>
    /// <example>{ "AAMkADIxYTJiZDIz", "C:\temp\file.txt", 6000, "#microsoft.graph.fileAttachment", "This is content." }</example>
    public List<ResultObject> Data { get; private set; }

    /// <summary>
    /// Error details. Null when Success is true.
    /// </summary>
    /// <example>null</example>
    public Error Error { get; private set; }
}
