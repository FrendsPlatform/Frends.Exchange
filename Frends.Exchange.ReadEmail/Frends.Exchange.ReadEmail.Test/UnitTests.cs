using Azure.Identity;
using Frends.Exchange.ReadEmail.Definitions;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
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
    private static Connection _connection = new();
    private static Input _input = new();
    private static Options _options = new();
    private static readonly string _downloadDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Download\\");

    [TestInitialize]
    public void Setup()
    {
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
    }

    [TestCleanup]
    public async Task CleanUp()
    {
        if (Directory.Exists(_downloadDir)) Directory.Delete(_downloadDir, true);
        await UpdateMessageRead();
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
    public async Task ReadEmailTest_ReadAndDownload_Inbox_ClientCredentialsSecret()
    {
        _connection.AuthenticationProvider = AuthenticationProviders.ClientCredentialsSecret;
        _connection.ClientSecret = _clientSecret;
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

    private static GraphServiceClient CreateGraphServiceClient()
    {
        var options = new TokenCredentialOptions { AuthorityHost = AzureAuthorityHosts.AzurePublicCloud };
        var credentials = new UsernamePasswordCredential(_user, _password, _tenantID, _applicationID, options);
        return new GraphServiceClient(credentials);
    }

    // Update message back to unread
    private static async Task UpdateMessageRead()
    {
        var client = CreateGraphServiceClient();
        var requestBody = new Message { IsRead = false };
        await client.Me.Messages["AAMkADIxYTJiZDIzLTIyZDMtNDhhNy05YjE1LTY2NGRkNmRjZTNiNwBGAAAAAACTqlZRkDG0S6Jj-VUkGGnxBwBGg69sLcQZTZPbCQVRM7fFAAAAAAEMAABGg69sLcQZTZPbCQVRM7fFAAFJtxHfAAA="].PatchAsync(requestBody);
    }
}