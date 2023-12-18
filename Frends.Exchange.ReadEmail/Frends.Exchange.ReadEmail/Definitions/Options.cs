using System.ComponentModel;

namespace Frends.Exchange.ReadEmail.Definitions;

/// <summary>
/// Options for controlling the behavior of a task.
/// </summary>
public class Options
{
    /// <summary>
    /// Gets or sets a value indicating whether an error should stop the task and throw an exception.
    /// If set to true, an exception will be thrown when an error occurs. If set to false, Task will try to continue and the error message will be added into Result.ErrorMessages and Result.Success will be set to false.
    /// </summary>
    /// <example>true</example>
    [DefaultValue(true)]
    public bool ThrowExceptionOnFailure { get; set; }
}