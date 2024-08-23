using Frends.Exchange.SendEmail.Definitions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Frends.Exchange.SendEmail.Tests;

[TestClass]
public class UnitTests
{
    private readonly string? _user = Environment.GetEnvironmentVariable("Exchange_User");
    private readonly string? _user2 = Environment.GetEnvironmentVariable("Exchange_User2");
    private readonly string? _password = Environment.GetEnvironmentVariable("Exchange_User_Password");
    private readonly string? _applicationID = Environment.GetEnvironmentVariable("Exchange_Application_ID");
    private readonly string? _tenantID = Environment.GetEnvironmentVariable("Exchange_Tenant_ID");
    private static Connection _connection = new();
    private static Input _input = new();
    private static Options _options = new();
    private static readonly string _filePath1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../first.txt");
    private static readonly string _filePath2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../second.txt");
    private static readonly string _largeFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../TheBigOne.txt");
    private static readonly string _largeFilePath2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../SecondTheBigOne.txt");
    private static readonly string[] _files = new[] { _filePath1, _filePath2, _largeFilePath, _largeFilePath2 };

    [TestInitialize]
    public void Setup()
    {
        CreateFiles();

        _connection = new Connection()
        {
            Username = _user,
            Password = _password,
            ClientId = _applicationID,
            TenantId = _tenantID,
            AuthenticationProvider = AuthenticationProviders.UsernamePassword,
            ClientSecret = null,
            X509CertificateFilePath = null,
        };

        _input = new Input()
        {
            From = _user,
            To = _user,
            Cc = null,
            Bcc = null,
            Message = $"This is a test message from Frends.Exchange.SendEmail Unit Tests sent in {DateTime.UtcNow}",
            IsMessageHtml = false,
            Subject = "This is subject",
            Attachments = null,
            Importance = ImportanceLevels.Normal,
            SaveToSentItems = true,
        };

        _options = new Options()
        {
            ThrowExceptionIfAttachmentNotFound = true,
            ThrowExceptionOnFailure = true,
        };
    }

    [TestCleanup]
    public void CleanUp()
    {
        DeleteFiles();
    }

    [TestMethod]
    public async Task SendEmailTest_PlainText_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_PlainText_UsernamePassword";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual("Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_ToMultiple_Semicolon_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_ToMultiple_Semicolon_UsernamePassword";
        _input.To = $"{_user}; {_user2}";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual("Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_ToMultiple_Comma_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_ToMultiple_Comma_UsernamePassword";
        _input.To = $"{_user}; {_user2}";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual("Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_WithCC_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_WithCC_UsernamePassword";
        _input.Cc = $"{_user}; {_user2}";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual("Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_WithBCC_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_WithBCC_UsernamePassword";
        _input.Bcc = $"{_user}; {_user2}";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual("Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_Html_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_Html_UsernamePassword";
        _input.Message = "<div><h1>This is a header text.</h1></div>";
        _input.IsMessageHtml = true;
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual("Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_NordicLetters_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_NordicLetters_UsernamePassword";
        _input.Message = "Tämä testimaili tuo yöllä ålannista.";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual("Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_SingleFileAttachment_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_SingleFileAttachment_UsernamePassword";
        _input.Attachments = new[] {
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = _filePath1,
                FileMask = "*.*"
            }
        };

        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual($"Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_MultipleFileAttachment_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_MultipleFileAttachment_UsernamePassword";
        _input.Attachments = new[] {
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = _filePath1,
                FileMask = "*.*"
            },
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = _filePath2,
                FileMask = "*.*"
            }
        };

        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual($"Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_SingleFileFromDir_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_SingleFileFromDir_UsernamePassword";
        _input.Attachments = new[] {
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = Path.GetDirectoryName(_filePath1),
                FileMask = $"{Path.GetFileNameWithoutExtension(_filePath1)}.*"
            }
        };

        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual($"Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_LargeAttachment_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_LargeAttachment_UsernamePassword";
        _input.From = null;
        _input.Attachments = new[] {
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = _largeFilePath,
                FileMask = "*.*"
            }
        };

        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual($"Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_MultipleLargeAttachments_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_MultipleLargeAttachments_UsernamePassword";
        _input.From = null;
        _input.Attachments = new[] {
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = _largeFilePath,
                FileMask = "*.*"
            },
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = _largeFilePath2,
                FileMask = "*.*"
            }
        };

        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual($"Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_StringAttachment_UsernamePassword()
    {

        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_StringAttachment_UsernamePassword";
        _input.Attachments = new[] {
            new Attachments()
            {
                AttachmentType = AttachmentTypes.AttachmentFromString,
                FileContent = "This is a test attachment from string.",
                FileName = "stringAttachmentFile.txt"
            }
        };

        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual($"Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_MultipleStringAttachment_UsernamePassword()
    {

        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_MultipleStringAttachment_UsernamePassword";
        _input.Attachments = new[] {
            new Attachments()
            {
                AttachmentType = AttachmentTypes.AttachmentFromString,
                FileContent = "This is a test attachment from string.",
                FileName = "stringAttachmentFile.txt"
            },
            new Attachments()
            {
                AttachmentType = AttachmentTypes.AttachmentFromString,
                FileContent = "This is another test attachment from string.",
                FileName = "stringAttachmentFile2.txt"
            }
        };

        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual($"Email sent successfully.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_MixAttachments_UsernamePassword()
    {

        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_MixAttachments_UsernamePassword";
        _input.Attachments = new[] {
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = Path.GetDirectoryName(_filePath1),
                FileMask = "*.*"
            },
            new Attachments()
            {
                AttachmentType = AttachmentTypes.AttachmentFromString,
                FileContent = "This is a test attachment from string.",
                FileName = "stringAttachmentFile.txt"
            }
        };

        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual($"Email sent successfully.", result.Data);
    }

    [TestMethod]
    [Ignore]
    /*
    This test cannot currently be run seemingly due to the test user not having the permission to "Send as" another user.
    */
    public async Task SendEmailTest_SendEmailAsAnotherUser_UsernamePassword()
    {
        _input.From = _user2;
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_SendEmailAsAnotherUser_UsernamePassword";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
    }

    [TestMethod]
    public async Task SendEmailTest_MissingTo_Throw()
    {
        _input.To = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.SendEmail(_connection, _input, _options, default));
    }

    [TestMethod]
    public async Task SendEmailTest_MissingCredentials_UsernamePassword_Throw()
    {
        _connection.TenantId = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.SendEmail(_connection, _input, _options, default));

        _connection.TenantId = _tenantID;
        _connection.ClientId = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.SendEmail(_connection, _input, _options, default));

        _connection.ClientId = _applicationID;
        _connection.Username = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.SendEmail(_connection, _input, _options, default));

        _connection.Username = _user;
        _connection.Password = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.SendEmail(_connection, _input, _options, default));
    }

    [TestMethod]
    public async Task SendEmailTest_MissingCredentials_ClientCredentialsCertificate_Throw()
    {
        _connection.AuthenticationProvider = AuthenticationProviders.ClientCredentialsCertificate;

        _connection.X509CertificateFilePath = "Something";
        _connection.TenantId = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.SendEmail(_connection, _input, _options, default));

        _connection.TenantId = _tenantID;
        _connection.ClientId = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.SendEmail(_connection, _input, _options, default));

        _connection.ClientId = _applicationID;
        _connection.X509CertificateFilePath = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.SendEmail(_connection, _input, _options, default));
    }

    [TestMethod]
    public async Task SendEmailTest_MissingCredentials_ClientCredentialsSecret_Throw()
    {
        _connection.AuthenticationProvider = AuthenticationProviders.ClientCredentialsSecret;

        _connection.ClientSecret = "Something";
        _connection.TenantId = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.SendEmail(_connection, _input, _options, default));

        _connection.TenantId = _tenantID;
        _connection.ClientId = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.SendEmail(_connection, _input, _options, default));

        _connection.ClientId = _applicationID;
        _connection.ClientSecret = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.SendEmail(_connection, _input, _options, default));
    }

    [TestMethod]
    public async Task SendEmailTest_FileNotFound()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_ReceiverNotFound";
        _input.Attachments = new[] {
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = @"C:\nothinghere",
                FileMask = "*.*"
            }
        };
        _options.ThrowExceptionOnFailure = false;
        _options.ThrowExceptionIfAttachmentNotFound = true;

        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsFalse(result.Success);
        Assert.AreEqual("Failed to send an email. No files found in directory C:\\nothinghere.", result.Data);
    }

    [TestMethod]
    public async Task SendEmailTest_FileNotFound2()
    {
        _input.Subject = $"{_input.Subject}, Method: SendEmailTest_ReceiverNotFound";
        _input.Attachments = new[] {
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = @"C:\nothinghere",
                FileMask = "*.*"
            }
        };
        _options.ThrowExceptionOnFailure = true;
        _options.ThrowExceptionIfAttachmentNotFound = true;
        await Assert.ThrowsExceptionAsync<Exception>(async () => await Exchange.SendEmail(_connection, _input, _options, default));
    }

    internal static void CreateFiles()
    {
        foreach (var filePath in _files)
            if (!File.Exists(filePath))
            {
                if (new FileInfo(filePath).Name.Contains("TheBigOne"))
                {
                    // Write 9MB file.
                    using var stream = new FileStream(filePath, FileMode.CreateNew);
                    stream.Seek(9 * 1024 * 1024, SeekOrigin.Begin);
                    stream.WriteByte(0);
                }
                else
                    File.WriteAllText(filePath, $"This is a test attachment {Path.GetFileName(filePath)}.");
            }
    }

    private static void DeleteFiles()
    {
        foreach (var filePath in _files)
            if (File.Exists(filePath))
                File.Delete(filePath);
    }
}