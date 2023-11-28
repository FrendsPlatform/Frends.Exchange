using Azure.Identity;
using Frends.Exchange.SendEmail.Definitions;
using Microsoft.Graph;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Frends.Exchange.SendEmail;

/// <summary>
/// Google Drive Download Task.
/// </summary>
public class Exchange
{
    /// <summary>
    /// Download objects from Google Drive.
    /// [Documentation](https://tasks.frends.com/tasks/frends-tasks/Frends.Exchange.SendEmail)
    /// </summary>
    /// <param name="connection">Connection parameters.</param>
    /// <param name="input">Input parameters</param>
    /// <param name="options">Options parameters.</param>
    /// <param name="cancellationToken">Token received from Frends to cancel this Task.</param>
    /// <returns>Object { bool Success, List&lt;Output&gt; Data }</returns>
    public static async Task<Result> SendEmail([PropertyTab] Connection connection, [PropertyTab] Input input, [PropertyTab] Options options, CancellationToken cancellationToken)
    {
        InputCheck(connection, input);
        using var client = CreateGraphServiceClient(connection);
        if (input.Attachments != null && input.Attachments.Length > 0)
            return new Result(true, await SendExchangeEmailWithAttachments(input, options, client, cancellationToken));
        else
            return new Result(true, await SendExchangeEmail(input, options.ThrowExceptionOnFailure, client, cancellationToken));
    }

    private static async Task<List<Output>> SendExchangeEmail(Input input, bool throwExceptionOnFailure,  GraphServiceClient graphClient, CancellationToken cancellationToken)
    {
        var outputList = new List<Output>();

        try
        {
            var message = GetMessageBody(input);
            var requestBody = new SendMailPostRequestBody()
            {
                Message = message,
                SaveToSentItems = input.SaveToSentItems
            };
            await graphClient.Me.SendMail.PostAsync(requestBody, cancellationToken: cancellationToken);
            outputList.Add(new Output() { EmailSent = true, MessageStatus = $"Email sent to {input.To}." });
            return outputList;
        }
        catch (Exception ex)
        {
            if (throwExceptionOnFailure)
                throw;
            outputList.Add(new Output() { EmailSent = false, MessageStatus = $"Failed to send an email to {input.To}. {ex}" });
        }

        return outputList;
    }

    private static async Task<List<Output>> SendExchangeEmailWithAttachments(Input input, Options options, GraphServiceClient graphClient, CancellationToken cancellationToken)
    {
        var outputList = new List<Output>();
        var message = GetMessageBody(input);
        var fileList = new List<(string filePath, bool isLarge)>(); // boolean is used to determine whether the file will be deleted after sending.

        foreach (var attachment in input.Attachments)
        {
            switch (attachment.AttachmentType)
            {
                case AttachmentTypes.FileAttachment:
                    // If the path ends in a directory, all files in that folder with given attachment.FileMask are added as attachments.
                    var files = Directory.Exists(attachment.FilePath) ? Directory.GetFiles(attachment.FilePath, attachment.FileMask) : new string[] { attachment.FilePath };

                    if (files.Length == 0 && options.ThrowExceptionIfAttachmentNotFound)
                        throw new Exception($"No files found in directory {attachment.FilePath}.");

                    foreach (var file in files)
                        fileList.Add((file, false));
                    break;
                case AttachmentTypes.AttachmentFromString:
                    var tempFilePath = CreateTemporaryFile(attachment);
                    fileList.Add((tempFilePath, true));
                    break;
            }
        }

        foreach (var (filePath, isLarge) in fileList)
        {
            var attachmentSizeInBytes = new FileInfo(filePath).Length;
            var fileName = Path.GetFileName(filePath);

            // 3MB or less
            if (attachmentSizeInBytes <= 3 * 1024 * 1024) 
            {
                byte[] littleStream = File.ReadAllBytes(filePath);
                message.Attachments = new List<Attachment>
                {
                    new Attachment
                    {
                        OdataType = "#microsoft.graph.fileAttachment",
                        Name = fileName,
                        AdditionalData = new Dictionary<string, object> { { "contentBytes", Convert.ToBase64String(littleStream) } },
                    }
                };

                try
                {
                    var postSmallAttachment = graphClient.Users[input.From].Messages.PostAsync(message, cancellationToken: cancellationToken).GetAwaiter().GetResult();
                    outputList.Add(new Output() { EmailSent = true, MessageStatus = $"Email sent to {input.To} with an attachment {fileName}" });
                }
                catch (Exception ex)
                {
                    if (options.ThrowExceptionOnFailure)
                        throw;
                    outputList.Add(new Output() { EmailSent = false, MessageStatus = $"Failed to send an email to {input.To} with an attachment {fileName}. {ex}" });
                }
            }
            // More than 3MB
            else
            {
                var maxSliceSize = 320 * 1024;
                var fileStream = File.OpenRead(filePath);
                var uploadRequestBody = new Microsoft.Graph.Users.Item.Messages.Item.Attachments.CreateUploadSession.CreateUploadSessionPostRequestBody
                {
                    AttachmentItem = new AttachmentItem
                    {
                        AttachmentType = AttachmentType.File,
                        Name = fileName,
                        Size = fileStream.Length,
                        ContentType = "application/octet-stream"
                    },
                };

                var uploadSession = graphClient.Users[input.From].Messages[message.Id].Attachments.CreateUploadSession.PostAsync(uploadRequestBody, cancellationToken: cancellationToken).GetAwaiter().GetResult();
                var fileUploadTask = new LargeFileUploadTask<FileAttachment>(uploadSession, fileStream, maxSliceSize);
                var totalLength = fileStream.Length;

                // Create a callback that is invoked after each slice is uploaded
                IProgress<long> progress = new Progress<long>(prog =>
                {
                    Console.WriteLine($"Uploaded {prog} bytes of {totalLength} bytes");
                });
                try
                {
                    var largeAttachmentUploadResult = fileUploadTask.UploadAsync(progress, cancellationToken: cancellationToken).GetAwaiter().GetResult();
                    if (largeAttachmentUploadResult.UploadSucceeded)
                        outputList.Add(new Output() { EmailSent = true, MessageStatus = $"Email sent to {input.To} with an attachment {fileName} (size: {fileStream.Length})." });
                    else
                    {
                        if (options.ThrowExceptionOnFailure)
                            throw new Exception($"Failed to send an email to {input.To} with an attachment {fileName}.");
                        outputList.Add(new Output() { EmailSent = false, MessageStatus = $"Failed to send an email to {input.To} with an attachment {fileName}." });
                    }
                }
                catch (Exception ex)
                {
                    if (options.ThrowExceptionOnFailure)
                        throw;
                    outputList.Add(new Output() { EmailSent = false, MessageStatus = ex.Message });
                }

                graphClient.Users[input.From].Messages[message.Id].Send.PostAsync(cancellationToken: cancellationToken).GetAwaiter().GetResult();
            }

            if (isLarge)
                CleanUpTempFiles(filePath);
        }
        return outputList;
    }

    private static void InputCheck(Connection connection, Input input)
    {
        switch (connection.AuthenticationProvider)
        {
            case AuthenticationProviders.ClientCredentialsCertificate:
                if (string.IsNullOrWhiteSpace(connection.TenantId) || string.IsNullOrWhiteSpace(connection.ClientId) || string.IsNullOrWhiteSpace(connection.X509CertificateFilePath))
                    throw new ArgumentException(@"One or more required connection values missing:", $"{nameof(connection.TenantId)}, {nameof(connection.ClientId)}, {nameof(connection.X509CertificateFilePath)}.");
                break;
            case AuthenticationProviders.ClientCredentialsSecret:
                if (string.IsNullOrWhiteSpace(connection.TenantId) || string.IsNullOrWhiteSpace(connection.ClientId) || string.IsNullOrWhiteSpace(connection.ClientSecret))
                    throw new ArgumentException(@"One or more required connection values missing:", $"{nameof(connection.TenantId)}, {nameof(connection.ClientId)}, {nameof(connection.ClientSecret)}.");
                break;
            case AuthenticationProviders.UsernamePassword:
                if (string.IsNullOrWhiteSpace(connection.Username) || string.IsNullOrWhiteSpace(connection.Password) || string.IsNullOrWhiteSpace(connection.TenantId) || string.IsNullOrWhiteSpace(connection.ClientId))
                    throw new ArgumentException(@"One or more required connection values missing:", $"{nameof(connection.Username)}, {nameof(connection.Password)}, {nameof(connection.TenantId)}, {nameof(connection.ClientId)}.");
                break;
        }

        if (string.IsNullOrWhiteSpace(input.From) || string.IsNullOrWhiteSpace(input.To))
            throw new ArgumentException(@"One or more required message values missing:", $"{nameof(input.From)}, {nameof(input.To)}");
    }

    private static GraphServiceClient CreateGraphServiceClient(Connection connection)
    {
        var options = new TokenCredentialOptions { AuthorityHost = AzureAuthorityHosts.AzurePublicCloud };

        switch (connection.AuthenticationProvider)
        {
            case AuthenticationProviders.ClientCredentialsCertificate:
                var clientCertCredential = new ClientCertificateCredential(connection.TenantId, connection.ClientId, connection.X509CertificateFilePath, options);
                return new GraphServiceClient(clientCertCredential);
            case AuthenticationProviders.ClientCredentialsSecret:
                var clientSecretCredential = new ClientSecretCredential(connection.TenantId, connection.ClientId, connection.ClientSecret, options);
                return new GraphServiceClient(clientSecretCredential);
            case AuthenticationProviders.UsernamePassword:
                var credentials = new UsernamePasswordCredential(connection.Username, connection.Password, connection.TenantId, connection.ClientId, options);
                return new GraphServiceClient(credentials);
            default:
                throw new ArgumentException($"Invalid {nameof(connection.AuthenticationProvider)}.");
        }
    }

    private static Message GetMessageBody(Input input)
    {
        return new Message
        {
            Subject = input.Subject,
            Body = new ItemBody
            {
                ContentType = input.IsMessageHtml ? BodyType.Html : BodyType.Text,
                Content = input.Message,
            },
            ToRecipients = string.IsNullOrWhiteSpace(input.To) ? null : GetRecipients(input.To),
            CcRecipients = string.IsNullOrWhiteSpace(input.Cc) ? null : GetRecipients(input.Cc),
            BccRecipients = string.IsNullOrWhiteSpace(input.Bcc) ? null : GetRecipients(input.Bcc),
            Importance = GetImportance(input.Importance),
        };
    }

    private static List<Recipient> GetRecipients(string to)
    {
        var recipientList = new List<Recipient>();
        var recipients = to.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var receiver in recipients)
            recipientList.Add(new Recipient { EmailAddress = new EmailAddress { Address = receiver } });

        return recipientList;
    }

    private static Importance GetImportance(ImportanceLevels importance)
    {
        return importance switch
        {
            ImportanceLevels.Low => Importance.Low,
            ImportanceLevels.Normal => Importance.Normal,
            ImportanceLevels.High => Importance.High,
            _ => throw new ArgumentException($"Invalid {nameof(importance)}."),
        };
    }

    private static string CreateTemporaryFile(Attachments attachments)
    {
        var tempWorkDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        Directory.CreateDirectory(tempWorkDir);
        var filePath = Path.Combine(tempWorkDir, attachments.FileName);
        var content = attachments.FileContent;
        using (var sw = File.CreateText(filePath)) sw.Write(content);
        return filePath;
    }

    private static void CleanUpTempFiles(string tempFile)
    {
        var dir = Path.GetDirectoryName(tempFile);
        if (File.Exists(tempFile)) File.Delete(tempFile);
        if (Directory.GetFiles(dir) is null) Directory.Delete(dir, true);
    }
}