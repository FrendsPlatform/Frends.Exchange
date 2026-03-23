using Azure.Identity;
using Frends.Exchange.ReadEmail.Definitions;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Frends.Exchange.ReadEmail.Tests;

[TestClass]
public class UnitTests
{

    private static readonly string? _user = Environment.GetEnvironmentVariable("Exchange_User");
    private static readonly string? _password = Environment.GetEnvironmentVariable("Exchange_User_Password");
    private static readonly string? _applicationID = Environment.GetEnvironmentVariable("Exchange_Application_ID");
    private static readonly string? _tenantID = Environment.GetEnvironmentVariable("Exchange_Tenant_ID");
    private static readonly string? _clientSecret = Environment.GetEnvironmentVariable("Exchange_ClientSecret");
    private const string TestTag = "FRENDS-TEST-MESSAGE";
    private static Connection _connection = new();
    private static Input _input = new();
    private static Options _options = new();
    private static readonly string _downloadDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Download\\");

    [TestInitialize]
    public async Task Setup()
    {
        _connection = new Connection()
        {
            Username = _user,
            Password = _password,
            ClientId = _applicationID,
            TenantId = _tenantID,
            AuthenticationProvider = AuthenticationProviders.ClientCredentialsSecret,
            ClientSecret = _clientSecret,
            X509CertificateFilePath = null,
        };

        _input = new Input()
        {
            Select = null,
            Filter = "parentFolderId eq 'INBOX'",
            Skip = 0,
            Top = 0,
            Orderby = null,
            Expand = null,
            Headers = null,
            DownloadAttachments = true,
            DestinationDirectory = _downloadDir,
            FileExistHandler = FileExistHandlers.Rename,
            CreateDirectory = true,
            From = _user,
            UpdateReadStatus = false,
        };

        _options = new Options()
        {
            ThrowExceptionOnFailure = true,
        };

        await SeedMailbox(3);

        await Task.Delay(5000);
    }

    [TestCleanup]
    public async Task CleanUp()
    {
        if (Directory.Exists(_downloadDir)) Directory.Delete(_downloadDir, true);
        await CleanUpTestEmails();
    }


    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_AllFolders()
    {
        _input.Filter = null;
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_Inbox()
    {
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
    }

    [TestMethod]
    public async Task ReadEmailTest_Read_DONTDownload_Inbox()
    {
        _input.DownloadAttachments = false;
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsFalse(Directory.Exists(_input.DestinationDirectory));
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_Select_Inbox()
    {
        _input.Select = "subject";
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_FROM_Inbox()
    {
        _input.From = _user;
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_MarkAsRead_Inbox()
    {
        _input.UpdateReadStatus = true;
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_Skip_Inbox()
    {
        _input.Skip = 2;
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_Top_Inbox()
    {
        _input.Top = 2;
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual(2, result.Data.Count);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_OrderBy_Inbox()
    {
        _input.Orderby = "receivedDateTime DESC,subject ASC";
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_Expand_Inbox()
    {
        _input.Expand = "attachments";
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_Header_Inbox()
    {
        _input.Headers = new[] { new HeaderParameters() { HeaderName = "Prefer", HeaderValues = new[] { "outlook.body-content-type=\"text\"" } } };
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_FileExistHandler_Skip()
    {
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
        var fileCount = Directory.GetFiles(_downloadDir).Length;

        _input.FileExistHandler = FileExistHandlers.Skip;
        var resultSkip = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(resultSkip.Success);
        Assert.IsTrue(resultSkip.Data.Count > 0);
        Assert.AreEqual(0, resultSkip.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
        Assert.AreEqual(fileCount, Directory.GetFiles(_downloadDir).Length);
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_FileExistHandler_Rename()
    {
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
        var fileCount = Directory.GetFiles(_downloadDir).Length;

        _input.FileExistHandler = FileExistHandlers.Rename;
        var resultSkip = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(resultSkip.Success);
        Assert.IsTrue(resultSkip.Data.Count > 0);
        Assert.AreEqual(0, resultSkip.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
        Assert.AreEqual(fileCount + fileCount, Directory.GetFiles(_downloadDir).Length);
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_FileExistHandler_OverWrite()
    {
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
        var fileCount = Directory.GetFiles(_downloadDir).Length;

        _input.FileExistHandler = FileExistHandlers.OverWrite;
        var resultSkip = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(resultSkip.Success);
        Assert.IsTrue(resultSkip.Data.Count > 0);
        Assert.AreEqual(0, resultSkip.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
        Assert.AreEqual(fileCount, Directory.GetFiles(_downloadDir).Length);
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDownload_FileExistHandler_Append()
    {
        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.Data.Count > 0);
        Assert.AreEqual(0, result.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
        var fileCount = Directory.GetFiles(_downloadDir).Length;

        _input.FileExistHandler = FileExistHandlers.Append;
        var resultSkip = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(resultSkip.Success);
        Assert.IsTrue(resultSkip.Data.Count > 0);
        Assert.AreEqual(0, resultSkip.ErrorMessages.Count);
        Assert.IsTrue(Directory.Exists(_input.DestinationDirectory));
        Assert.AreEqual(fileCount, Directory.GetFiles(_downloadDir).Length);
    }

    [TestMethod]
    public async Task ReadEmailTest_TryToDownload_NoDir()
    {
        _input.CreateDirectory = false;
        await Assert.ThrowsExceptionAsync<Exception>(async () => await Exchange.ReadEmail(_connection, _input, _options, default));
    }

    [TestMethod]
    public async Task ReadEmailTest_MissingCredentials_UsernamePassword_Throw()
    {
        _connection.AuthenticationProvider = AuthenticationProviders.UsernamePassword;

        _connection.TenantId = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.ReadEmail(_connection, _input, _options, default));

        _connection.TenantId = _tenantID;
        _connection.ClientId = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.ReadEmail(_connection, _input, _options, default));

        _connection.ClientId = _applicationID;
        _connection.Username = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.ReadEmail(_connection, _input, _options, default));

        _connection.Username = _user;
        _connection.Password = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.ReadEmail(_connection, _input, _options, default));
    }

    [TestMethod]
    public async Task ReadEmailTest_MissingCredentials_ClientCredentialsCertificate_Throw()
    {
        _connection.AuthenticationProvider = AuthenticationProviders.ClientCredentialsCertificate;

        _connection.X509CertificateFilePath = "Something";
        _connection.TenantId = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.ReadEmail(_connection, _input, _options, default));

        _connection.TenantId = _tenantID;
        _connection.ClientId = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.ReadEmail(_connection, _input, _options, default));

        _connection.ClientId = _applicationID;
        _connection.X509CertificateFilePath = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.ReadEmail(_connection, _input, _options, default));
    }

    [TestMethod]
    public async Task ReadEmailTest_MissingCredentials_ClientCredentialsSecret_Throw()
    {
        _connection.AuthenticationProvider = AuthenticationProviders.ClientCredentialsSecret;

        _connection.ClientSecret = "Something";
        _connection.TenantId = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.ReadEmail(_connection, _input, _options, default));

        _connection.TenantId = _tenantID;
        _connection.ClientId = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.ReadEmail(_connection, _input, _options, default));

        _connection.ClientId = _applicationID;
        _connection.ClientSecret = null;
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () => await Exchange.ReadEmail(_connection, _input, _options, default));
    }

    [TestMethod]
    public async Task ReadEmailTest_ReadAndDelete()
    {
        var testSubject = $"TEST DELETE EMAIL {Guid.NewGuid()}";
        await SendTestEmail(testSubject);

        await Task.Delay(5000);

        _input.Filter = $"parentFolderId eq 'INBOX' and subject eq '{testSubject}'";
        _options.DeleteReadEmails = true;

        var result = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        Assert.AreEqual(1, result.Data.Count, "Should have found exactly one test email");
        Assert.AreEqual(testSubject, result.Data[0].Subject);
        Assert.AreEqual(0, result.ErrorMessages.Count);

        _options.DeleteReadEmails = false;
        var secondRead = await Exchange.ReadEmail(_connection, _input, _options, default);
        Assert.AreEqual(0, secondRead.Data.Count, "Test email should have been deleted");
    }

    [TestMethod]
    public async Task ReadEmailTest_NoThrow_ReturnErrorList()
    {
        // Force a failure by providing an invalid TenantId
        _connection.TenantId = "invalid-guid";
        _options.ThrowExceptionOnFailure = false;

        var result = await Exchange.ReadEmail(_connection, _input, _options, default);

        // This checks the 'else' branch in your catch block
        Assert.IsFalse(result.Success);
        Assert.IsTrue(result.ErrorMessages.Count > 0);
        Assert.AreEqual(0, result.Data.Count);
    }

    private static GraphServiceClient CreateGraphServiceClient()
    {
        var options = new TokenCredentialOptions { AuthorityHost = AzureAuthorityHosts.AzurePublicCloud };
        var credentials = new ClientSecretCredential(_tenantID, _applicationID, _clientSecret, options);
        return new GraphServiceClient(credentials);
    }

    private async Task SendTestEmail(string subject)
    {
        var client = CreateGraphServiceClient();

        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                ContentType = BodyType.Text,
                Content = "This is a test email with an attachment to ensure the directory is created."
            },
            ToRecipients = new List<Recipient>
        {
            new Recipient
            {
                EmailAddress = new EmailAddress { Address = _user }
            }
        },
            Attachments = new List<Attachment>
        {
            new FileAttachment
            {
                OdataType = "#microsoft.graph.fileAttachment",
                Name = "test-attachment.txt",
                ContentType = "text/plain",
                ContentBytes = System.Text.Encoding.UTF8.GetBytes("Hello World! This is a test attachment.")
            }
        }
        };

        var requestBody = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
        {
            Message = message,
            SaveToSentItems = false
        };

        await client.Users[_user].SendMail.PostAsync(requestBody);
    }

    private async Task SeedMailbox(int count = 1)
    {
        for (int i = 0; i < count; i++)
        {
            await SendTestEmail($"{TestTag} {Guid.NewGuid()}");
        }
        await Task.Delay(2000);
    }

    private async Task CleanUpTestEmails()
    {
        var client = CreateGraphServiceClient();

        var messages = await client.Users[_user].Messages
            .GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Filter = $"contains(subject, '{TestTag}')";
            });

        if (messages?.Value != null)
        {
            foreach (var msg in messages.Value)
            {
                await client.Users[_user].Messages[msg.Id].DeleteAsync();
            }
        }
    }
}