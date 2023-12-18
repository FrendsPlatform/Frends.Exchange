using System.Collections.Generic;

namespace Frends.Exchange.ReadEmail.Definitions;

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
    /// <example>{ "AAMkADIxYTJiZDIz", "C:\temp\file.txt", 6000, "#microsoft.graph.fileAttachment", "This is content." }</example>
    public List<ResultObject> Data { get; private set; }

    /// <summary>
    /// Error messages.
    /// </summary>
    /// <example>{ "error occured", "another error occured." }</example>
    public List<dynamic> ErrorMessages { get; private set; }

    internal Result(bool success, List<ResultObject> data, List<dynamic> errorMessage)
    {
        Success = success;
        Data = data;
        ErrorMessages = errorMessage;
    }
}
