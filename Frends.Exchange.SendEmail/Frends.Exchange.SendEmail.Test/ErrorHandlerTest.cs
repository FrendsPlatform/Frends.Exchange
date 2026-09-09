using Frends.Exchange.SendEmail.Definitions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading.Tasks;

namespace Frends.Exchange.SendEmail.Tests;

[TestClass]
public class ErrorHandlerTest
{
    private const string CustomErrorMessage = "CustomErrorMessage";
    private static Connection _connection = new();
    private static Input _input = new();
    private static Options _options = new();

    [TestInitialize]
    public void Setup()
    {
        _connection = new Connection()
        {
            Username = "test",
            Password = "test",
            ClientId = "test",
            TenantId = "test",
            AuthenticationProvider = AuthenticationProviders.UsernamePassword,
            ClientSecret = null,
            X509CertificateFilePath = null,
        };

        _input = new Input()
        {
            From = "test",
            To = "test",
            Subject = "This is subject",
            Message = "This is message",
            Attachments = new[]
            {
                new Attachments()
                {
                    AttachmentType = AttachmentTypes.FileAttachment,
                    FilePath = @"C:\nothinghere",
                    FileMask = "*.*",
                },
            },
        };

        _options = new Options()
        {
            ThrowErrorOnFailure = true,
            ThrowExceptionIfAttachmentNotFound = true,
        };
    }

    [TestMethod]
    public async Task Should_Throw_Error_When_ThrowErrorOnFailure_Is_True()
    {
        var ex = await Assert.ThrowsExceptionAsync<Exception>(async () =>
            await Exchange.SendEmail(_input, _connection, _options, default));
        Assert.IsNotNull(ex);
    }

    [TestMethod]
    public async Task Should_Return_Failed_Result_When_ThrowErrorOnFailure_Is_False()
    {
        _options.ThrowErrorOnFailure = false;
        var result = await Exchange.SendEmail(_input, _connection, _options, default);
        Assert.IsFalse(result.Success);
    }

    [TestMethod]
    public async Task Should_Use_Custom_ErrorMessageOnFailure()
    {
        _options.ErrorMessageOnFailure = CustomErrorMessage;
        var ex = await Assert.ThrowsExceptionAsync<Exception>(async () =>
            await Exchange.SendEmail(_input, _connection, _options, default));
        Assert.IsNotNull(ex);
        StringAssert.Contains(ex.Message, CustomErrorMessage);
    }
}
