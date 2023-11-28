using Azure.Identity;
using Frends.Exchange.SendEmail.Definitions;
using Microsoft.Graph.Models;
using Microsoft.Graph;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using Microsoft.VisualStudio.TestPlatform.CommunicationUtilities;
using Microsoft.Kiota.Abstractions;
using Microsoft.Graph.Models.Security;

namespace Frends.Exchange.SendEmail.Tests;

[TestClass]
public class UnitTests
{
    private readonly string? _user = Environment.GetEnvironmentVariable("Exchange_User");
    private readonly string? _user2 = Environment.GetEnvironmentVariable("Exchange_User2");
    private readonly string? _password = Environment.GetEnvironmentVariable("Exchange_User_Password");
    private readonly string? _applicationID = Environment.GetEnvironmentVariable("Exchange_Application_ID");
    private readonly string? _tenantID = Environment.GetEnvironmentVariable("Exchange_Tenant_ID");
    private string _username = string.Empty;
    private static Connection _connection = new();
    private static Input _input = new();
    private static Options _options = new();

    [TestInitialize]
    public void Setup()
    {
        if (string.IsNullOrEmpty(_applicationID) || string.IsNullOrEmpty(_password) || string.IsNullOrEmpty(_tenantID))
            throw new ArgumentException("Password, Application ID or Tenant ID is missing. Please check environment variables.");

        _username = $"{_user}@frends.com";

        _connection = new Connection()
        {
            Username = "frends_exchange_test_user@hiq.fi", //_username,
            Password = _password,
            ClientId = _applicationID,
            TenantId = _tenantID,
            AuthenticationProvider = AuthenticationProviders.UsernamePassword,
            ClientSecret = null,
            X509CertificateFilePath = null,
        };

        _input = new Input()
        {
            From = "frends_exchange_test_user@hiq.fi",//_username,
            To = "frends_exchange_test_user@hiq.fi", //_username,
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

    }

    [TestMethod]
    public async Task SendEmailWithPlainTextTest_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: {System.Reflection.MethodBase.GetCurrentMethod()?.Name}";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        //var email = await ReadTestEmail();
        //Assert.AreEqual(_input.Message, email[0].BodyText);
        //await DeleteMessages();
    }

    /*
    [TestMethod]
    public async Task SendEmailToMultipleUsingSemicolonTest_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: {System.Reflection.MethodBase.GetCurrentMethod()?.Name}";
        _input.To = $"{_username}; {_username}";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        var email = await ReadTestEmail();
        Assert.AreEqual($"{_username}; {_username}", email[0].To);
        await DeleteMessages();
    }

    [TestMethod]
    public async Task SendEmailToMultipleUsingCommaTest_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: {System.Reflection.MethodBase.GetCurrentMethod()?.Name}";
        _input.To = $"{_username}, {_username}";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        var email = await ReadTestEmail();
        Assert.AreEqual($"{_username}, {_username}", email[0].To);
        await DeleteMessages();
    }

    [TestMethod]
    public async Task SendEmailWithCCTest_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: {System.Reflection.MethodBase.GetCurrentMethod()?.Name}";
        _input.Cc = $"{_username}, {_username}";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        var email = await ReadTestEmail();
        Assert.AreEqual($"{_username}, {_username}", email[0].Cc);
        await DeleteMessages();
    }

    [TestMethod]
    public async Task SendEmailWithBCCTest_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: {System.Reflection.MethodBase.GetCurrentMethod()?.Name}";
        _input.Bcc = $"{_username}, {_username}";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        var email = await ReadTestEmail();
        Assert.AreEqual($"{_username}, {_username}", email[0].Bcc);
        await DeleteMessages();
    }

    [TestMethod]
    public async Task SendEmailWithHtmlTest_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: {System.Reflection.MethodBase.GetCurrentMethod()?.Name}";
        _input.Message = "<div><h1>This is a header text.</h1></div>";
        _input.IsMessageHtml = true;
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        var email = await ReadTestEmail();
        Assert.IsTrue(email[0].BodyHtml.Contains(_input.Message));
        await DeleteMessages();
    }

    [TestMethod]
    public async Task SendEmailNordicLettersTest_UsernamePassword()
    {
        _input.Subject = $"{_input.Subject}, Method: {System.Reflection.MethodBase.GetCurrentMethod()?.Name}";
        _input.Message = "Tämä testimaili tuo yöllä ålannista.";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        var email = await ReadTestEmail();
        Assert.AreEqual("Tämä testimaili tuo yöllä ålannista.", email[0].BodyText);
        Assert.AreEqual(_input.Subject, email[0].Subject);
        await DeleteMessages();
    }

    [TestMethod]
    public async Task SendEmailWithFileAttachmentTest_UsernamePassword()
    {
        var filePath1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../first.txt");
        CreateFiles(new[] { filePath1 }, false);

        _input.Subject = $"{_input.Subject}, Method: {System.Reflection.MethodBase.GetCurrentMethod()?.Name}";
        _input.Attachments = new[] {
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = filePath1,
                FileMask = "*.*"
            }
        };

        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        DeleteFiles(new[] { filePath1 });
        await ReadTestEmailWithAttachment(_input.Subject);
        Assert.IsTrue(File.Exists(filePath1));
        DeleteFiles(new[] { filePath1 });
        await DeleteMessages();
    }

    [TestMethod]
    public async Task SendEmailWithMultipleFileAttachmentTest_UsernamePassword()
    {
        var filePath1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../first.txt");
        var filePath2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../second.txt");
        CreateFiles(new[] { filePath1, filePath2 }, false);
        
        _input.Subject = $"{_input.Subject}, Method: {System.Reflection.MethodBase.GetCurrentMethod()?.Name}";
        _input.Attachments = new[] {
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = filePath1,
                FileMask = "*.*"
            },
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = filePath2,
                FileMask = "*.*"
            }
        };

        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        DeleteFiles(new[] { filePath1, filePath2 });
        await ReadTestEmailWithAttachment(_input.Subject);
        Assert.IsTrue(File.Exists(filePath1));
        Assert.IsTrue(File.Exists(filePath2));
        DeleteFiles(new[] { filePath1, filePath2 });
        await DeleteMessages();
    }

    [TestMethod]
    public async Task SendEmailWithBigFileAttachmentTest_UsernamePassword()
    {
        var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../BigAttachmentFile.txt");
        CreateFiles(new[] { filePath }, true);

        _input.Subject = $"{_input.Subject}, Method: {System.Reflection.MethodBase.GetCurrentMethod()?.Name}";
        _input.Attachments = new[] {
            new Attachments()
            {
                AttachmentType = AttachmentTypes.FileAttachment,
                FilePath = filePath,
                FileMask = "*.*"
            }
        };

        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        DeleteFiles(new[] { filePath });
        await ReadTestEmailWithAttachment(_input.Subject);
        Assert.IsTrue(File.Exists(filePath));
        DeleteFiles(new[] { filePath });
        await DeleteMessages();
    }

    [TestMethod]
    public async Task SendEmailWithStringAttachmentTest_UsernamePassword()
    {
        var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../stringAttachmentFile.txt");
        CreateFiles(new[] { filePath }, false);

        _input.Subject = $"{_input.Subject}, Method: {System.Reflection.MethodBase.GetCurrentMethod()?.Name}";
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
        DeleteFiles(new[] { filePath });
        await ReadTestEmailWithAttachment(_input.Subject);
        Assert.IsTrue(File.Exists(filePath));
        DeleteFiles(new[] { filePath });
        await DeleteMessages();
    }

    [TestMethod]
    public async Task SendEmailAsAnotherUserTest_UsernamePassword()
    {
        _input.From = _user2;
        _input.Subject = $"{_input.Subject}, Method: {System.Reflection.MethodBase.GetCurrentMethod()?.Name}";
        var result = await Exchange.SendEmail(_connection, _input, _options, default);
        Assert.IsTrue(result.Success);
        var email = await ReadTestEmail();
        Assert.AreEqual("frends_exchange_test_user_2@frends.com", email[0].From);
        await DeleteMessages();
    }
    */

    //private async Task DeleteMessages()
    //{
    //    Thread.Sleep(5000); // Give some time for emails to get through before deletion.
    //    var options = new List<QueryOption>
    //        {
    //            new QueryOption("$search", "\"subject:" + _input.Subject + "\"")
    //        };
    //    var credentials = new UsernamePasswordCredential(_username, _password, _tenantID, _applicationID);
    //    var graph = new GraphServiceClient(credentials);
    //    var messages = await graph.Me.Messages.Request().GetAsync();
    //    foreach (var message in messages)
    //        await graph.Me.Messages[message.Id].Request().DeleteAsync();
    //}

    //private async Task<List<EmailMessageResult>> ReadTestEmail()
    //{
    //    Thread.Sleep(2000);

    //    var settings = new ExchangeSettings
    //    {
    //        TenantId = _tenantID,
    //        AppId = _applicationID,
    //        Username = _username,
    //        Password = _password
    //    };

    //    var options = new ExchangeOptions
    //    {
    //        MaxEmails = 1,
    //        DeleteReadEmails = false,
    //        GetOnlyUnreadEmails = false,
    //        MarkEmailsAsRead = false,
    //        IgnoreAttachments = true,
    //        EmailSubjectFilter = _input.Subject
    //    };

    //    var result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
    //    return result;
    //}

    //private async Task<List<EmailMessageResult>> ReadTestEmailWithAttachment(string subject)
    //{
    //    var dirPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../..");
    //    var settings = new ExchangeSettings
    //    {
    //        TenantId = _tenantID,
    //        AppId = _applicationID,
    //        Username = _username,
    //        Password = _password
    //    };

    //    var options = new ExchangeOptions
    //    {
    //        MaxEmails = 1,
    //        DeleteReadEmails = false,
    //        GetOnlyUnreadEmails = false,
    //        MarkEmailsAsRead = false,
    //        IgnoreAttachments = false,
    //        AttachmentSaveDirectory = dirPath,
    //        EmailSubjectFilter = subject
    //    };

    //    var result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
    //    return result;
    //}

    private static void CreateFiles(string[] filePaths, bool isBig)
    {
        foreach(var filePath in filePaths)
            if (!isBig)
                File.WriteAllText(filePath, $"This is a test attachment {Path.GetFileName(filePath)}.");
            else
            {
                // Write 9MB file.
                using var stream = new FileStream(filePath, FileMode.CreateNew);
                stream.Seek(9 * 1024 * 1024, SeekOrigin.Begin);
                stream.WriteByte(0);
            }
    }

    private static void DeleteFiles(string[] filePaths)
    {
        foreach(var filePath in filePaths) 
            File.Delete(filePath);
    }
}