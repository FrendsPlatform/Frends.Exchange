using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Exchange.ReadEmail.Definitions;

/// <summary>
/// Represents the parameters required to establish a connection.
/// </summary>
public class Connection
{
    /// <summary>
    /// Specifies the type of authentication provider to use for the connection.
    /// </summary>
    /// <remarks>
    /// This property determines the type of authentication to use when establishing a connection. The available options are:
    /// - `AuthenticationProviders.UsernamePassword`: This option uses a username and password for authentication.
    /// - `AuthenticationProviders.ClientCredentialsCertificate`: This option uses a client certificate and private key for authentication.
    /// - `AuthenticationProviders.ClientCredentialsSecret`: This option uses a client secret for authentication.
    /// </remarks>
    /// <example>
    /// AuthenticationProviders.UsernamePassword
    /// </example>
    [DefaultValue(AuthenticationProviders.UsernamePassword)]
    public AuthenticationProviders AuthenticationProvider { get; set; }

    /// <summary>
    /// Specifies the path to a file that contains both the client certificate and private key.
    /// </summary>
    /// <example>
    /// C:\Certificates\mycert.pfx
    /// </example>
    [UIHint(nameof(AuthenticationProvider), "", AuthenticationProviders.ClientCredentialsCertificate)]
    public string X509CertificateFilePath { get; set; }

    /// <summary>
    /// Specifies the secret key of the client application. 
    /// </summary>
    /// <example>
    /// Y2lzY29zeXN0ZW1zOmMxc2Nv
    /// </example>
    [UIHint(nameof(AuthenticationProvider), "", AuthenticationProviders.ClientCredentialsSecret)]
    [DisplayFormat(DataFormatString = "Text")]
    [PasswordPropertyText]
    public string ClientSecret { get; set; }

    /// <summary>
    /// Specifies the username of the user. 
    /// </summary>
    /// <example>
    /// johndoe@example.com
    /// </example>
    [UIHint(nameof(AuthenticationProvider), "", AuthenticationProviders.UsernamePassword)]
    public string Username { get; set; }

    /// <summary>
    /// Specifies the password of the user.
    /// </summary>
    /// <example>
    /// SecurePassword123
    /// </example>
    [UIHint(nameof(AuthenticationProvider), "", AuthenticationProviders.UsernamePassword)]
    [DisplayFormat(DataFormatString = "Text")]
    [PasswordPropertyText]
    public string Password { get; set; }

    /// <summary>
    /// Specifies the app ID for fetching an access token. 
    /// This is the unique identifier for your application.
    /// </summary>
    /// <example>
    /// 4a1aa1d9-c16a-40a2-bd7d-2bd40babe4ff
    /// </example>
    public string ClientId { get; set; }

    /// <summary>
    /// Specifies the tenant ID for fetching an access token. 
    /// This is the unique identifier of your Azure AD tenant.
    /// </summary>
    /// <example>
    /// 9188040d-6c67-4c5b-b112-36a304b66dad
    /// </example>
    public string TenantId { get; set; }
}