namespace Frends.Exchange.ReadEmail.Definitions;

/// <summary>
/// Represents the parameters for a header.
/// </summary>
public class HeaderParameters
{
    /// <summary>
    /// Specifies the name of the header to which values will be added.
    /// </summary>
    /// <example>Prefer</example>
    public string HeaderName { get; set; }

    /// <summary>
    /// Specifies the values to add to the header.
    /// </summary>
    /// <example>["outlook.body-content-type=\"text\", "foo""]</example>
    public string[] HeaderValues { get; set; }
}