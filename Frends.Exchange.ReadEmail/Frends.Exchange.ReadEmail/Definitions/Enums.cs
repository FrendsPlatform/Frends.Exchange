namespace Frends.Exchange.ReadEmail.Definitions;

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

/// <summary>
/// Specifies the action to take when a file already exists.
/// </summary>
public enum FileExistHandlers
{
    /// <summary>
    /// Skip the file creation.
    /// </summary>
    Skip,
    /// <summary>
    /// Rename the file by appending a unique number.
    /// </summary>
    Rename,
    /// <summary>
    /// Append to the existing file.
    /// </summary>
    Append,
    /// <summary>
    /// Overwrite the existing file.
    /// </summary>
    OverWrite
}