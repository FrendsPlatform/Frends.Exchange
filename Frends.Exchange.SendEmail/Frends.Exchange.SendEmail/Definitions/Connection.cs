using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Exchange.SendEmail.Definitions;

/// <summary>
/// Parameters for establishing a connection.
/// </summary>
public class Connection
{
    /// <summary>
    /// Specifies the authentication provider to use.
    /// </summary>
    /// <remarks>
    /// This determines the type of authentication to use when establishing a connection. The available options are:
    /// - `AuthenticationProviders.UsernamePassword`: Uses a username and password to authenticate.
    /// - `AuthenticationProviders.ClientCredentialsCertificate`: Uses a client certificate and private key to authenticate.
    /// - `AuthenticationProviders.ClientCredentialsSecret`: Uses a client secret to authenticate.
    /// </remarks>
    [DefaultValue(AuthenticationProviders.UsernamePassword)]
    public AuthenticationProviders AuthenticationProvider { get; set; }

    /// <summary>
    /// The path to a file that contains both the client certificate and private key.
    /// </summary>
    /// <example>c:\temp\mycert.pfx</example>
    [UIHint(nameof(AuthenticationProvider), "", AuthenticationProviders.ClientCredentialsCertificate)]
    public string X509CertificateFilePath { get; set; }

    /// <summary>
    /// The secret key of the client application. 
    /// </summary>
    /// <example>Y2lzY29zeXN0ZW1zOmMxc2Nv</example>
    [UIHint(nameof(AuthenticationProvider), "", AuthenticationProviders.ClientCredentialsSecret)]
    public string ClientSecret { get; set; }

    /// <summary>
    /// The username of the user. 
    /// </summary>
    /// <example>username@example.com</example>
    [UIHint(nameof(AuthenticationProvider), "", AuthenticationProviders.UsernamePassword)]
    public string Username { get; set; }

    /// <summary>
    /// The password of the user.
    /// </summary>
    /// <example>password123</example>
    [UIHint(nameof(AuthenticationProvider), "", AuthenticationProviders.UsernamePassword)]
    [PasswordPropertyText]
    public string Password { get; set; }

    /// <summary>
    /// The app ID for fetching an access token. 
    /// This is the unique identifier for your application.
    /// </summary>
    /// <example>4a1aa1d9-c16a-40a2-bd7d-2bd40babe4ff</example>
    [DefaultValue("")]
    public string ClientId { get; set; }

    /// <summary>
    /// The tenant ID for fetching an access token. 
    /// This is the unique identifier of your Azure AD tenant.
    /// </summary>
    /// <example>9188040d-6c67-4c5b-b112-36a304b66dad</example>
    [DefaultValue("")]
    public string TenantId { get; set; }
}