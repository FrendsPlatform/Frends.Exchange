using Azure.Identity;
using Frends.Exchange.ReadEmail.Definitions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading.Tasks;

namespace Frends.Exchange.ReadEmail.Tests;

[TestClass]
public class ErrorHandlerTest
{
    private const string CustomErrorMessage = "CustomErrorMessage";
    private static Connection _connection = new();
    private static Input _input = new();

    [TestMethod]
    public async Task Should_Throw_Error_When_ThrowErrorOnFailure_Is_True()
    {
        var options = new Options
        {
            ThrowErrorOnFailure = true
        };
        var ex = await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () =>
            await Exchange.ReadEmail(_input, _connection, options, default));
        Assert.IsNotNull(ex);
    }

    [TestMethod]
    public async Task Should_Return_Failed_Result_When_ThrowErrorOnFailure_Is_False()
    {
        var options = new Options
        {
            ThrowErrorOnFailure = false
        };
        var result = await Exchange.ReadEmail(_input, _connection, options, default);
        Assert.IsFalse(result.Success);
    }

    [TestMethod]
    public async Task Should_Use_Custom_ErrorMessageOnFailure()
    {
        var options = new Options
        {
            ThrowErrorOnFailure = true,
            ErrorMessageOnFailure = CustomErrorMessage
        };
        var ex = await Assert.ThrowsExceptionAsync<Exception>(async () =>
            await Exchange.ReadEmail(_input, _connection, options, default));
        Assert.IsNotNull(ex);
        StringAssert.Contains(ex.Message, CustomErrorMessage);
    }
}
