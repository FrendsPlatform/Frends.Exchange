namespace Frends.Exchange.SendEmail.Definitions;

/// <summary>
/// Attachment type options.
/// </summary>
public enum AttachmentTypes
{
    /// <summary>
    /// Select this if the attachment is a file. This means the attachment will be read from a file on disk.
    /// </summary>
    FileAttachment,

    /// <summary>
    /// Select this if the attachment file should be created from a string. This means the attachment will be created from a string in your code.
    /// </summary>
    AttachmentFromString
}

/// Importance level options.
public enum ImportanceLevels
{
    /// <summary>
    /// Select this if the email is of low importance. This might be used for routine or informational emails.
    /// </summary>
    Low,

    /// <summary>
    /// Select this if the email is of normal importance. This is the default value for most emails.
    /// </summary>
    Normal,

    /// <summary>
    /// Select this if the email is of high importance. This might be used for urgent or time-sensitive emails.
    /// </summary>
    High
}

/// <summary>
/// Authentication provider options.
/// </summary>
public enum AuthenticationProviders
{
    /// <summary>
    /// Select this if the authentication should be done using a client certificate. This requires a certificate file and private key.
    /// </summary>
    ClientCredentialsCertificate,

    /// <summary>
    /// Select this if the authentication should be done using a client secret. This requires a secret key from your application registration.
    /// </summary>
    ClientCredentialsSecret,

    /// <summary>
    /// Select this if the authentication should be done using a username and password. This requires the username and password of the user.
    /// </summary>
    UsernamePassword
}