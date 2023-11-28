using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Exchange.SendEmail.Definitions;

/// <summary>
/// Connection parameters.
/// </summary>
public class Connection
{
    public AuthenticationProviders AuthenticationProvider { get; set; }

    /// <summary>
    /// The path to a file which contains both the client certificate and private key.
    /// </summary>
    /// <example>c:\temp\mycert.pfx</example>
    [UIHint(nameof(AuthenticationProviders), "", AuthenticationProviders.ClientCredentialsCertificate)]
    public string X509CertificateFilePath { get; set; }

    /// <summary>
    /// The secret key of the client application. 
    /// </summary>
    /// <example>Y2lzY29zeXN0ZW1zOmMxc2Nv</example>
    [UIHint(nameof(AuthenticationProviders), "", AuthenticationProviders.ClientCredentialsSecret)]
    public string ClientSecret { get; set; }

    /// <summary>
    /// Username of the user. 
    /// </summary>
    /// <example>username@example.com</example>
    [UIHint(nameof(AuthenticationProviders), "", AuthenticationProviders.UsernamePassword)]
    public string Username { get; set; }

    /// <summary>
    /// Use this password to log in to the SMTP server.
    /// This is used along with the username to log in to the SMTP server.
    /// </summary>
    /// <example>password123</example>
    [UIHint(nameof(AuthenticationProviders), "", AuthenticationProviders.UsernamePassword)]
    [PasswordPropertyText]
    public string Password { get; set; }

    /// <summary>
    /// App ID for fetching access token. 
    /// This is the unique identifier for your application.
    /// </summary>
    /// <example>4a1aa1d9-c16a-40a2-bd7d-2bd40babe4ff</example>
    [DefaultValue("")]
    public string ClientId { get; set; }

    /// <summary>
    /// Tenant ID for fetching access token. 
    /// This is the unique identifier of your Azure AD tenant.
    /// </summary>
    /// <example>9188040d-6c67-4c5b-b112-36a304b66dad</example>
    [DefaultValue("")]
    public string TenantId { get; set; }
}