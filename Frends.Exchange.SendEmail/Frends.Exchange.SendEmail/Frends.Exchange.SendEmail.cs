using Azure.Identity;
using Frends.Exchange.SendEmail.Definitions;
using Microsoft.Graph;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Frends.Exchange.SendEmail;

/// <summary>
/// Microsoft Exchange Task.
/// </summary>
public class Exchange
{
    /// <summary>
    /// List of temp files to be deleted.
    /// </summary>
    internal static List<string> tempFilePaths = new();

    /// <summary>
    /// Send a Microsoft Exchange email.
    /// [Documentation](https://tasks.frends.com/tasks/frends-tasks/Frends.Exchange.SendEmail)
    /// </summary>
    /// <param name="connection">Parameters for establishing a connection.</param>
    /// <param name="input">Email content</param>
    /// <param name="options">Options for controlling the behavior of this Task.</param>
    /// <param name="cancellationToken">Token received from Frends to cancel this Task.</param>
    /// <returns>Object { bool Success, string Data }</returns>
    public static async Task<Result> SendEmail([PropertyTab] Connection connection, [PropertyTab] Input input, [PropertyTab] Options options, CancellationToken cancellationToken)
    {
        InputCheck(connection, input);
        return await SendExchangeEmail(input, connection, options, cancellationToken);
    }

    private static void InputCheck(Connection connection, Input input)
    {
        switch (connection.AuthenticationProvider)
        {
            case AuthenticationProviders.ClientCredentialsCertificate:
                if (string.IsNullOrWhiteSpace(connection.TenantId) || string.IsNullOrWhiteSpace(connection.ClientId) || string.IsNullOrWhiteSpace(connection.X509CertificateFilePath))
                    throw new ArgumentNullException(@"One or more required connection values missing:", $"{nameof(connection.TenantId)}, {nameof(connection.ClientId)}, {nameof(connection.X509CertificateFilePath)}.");
                break;
            case AuthenticationProviders.ClientCredentialsSecret:
                if (string.IsNullOrWhiteSpace(connection.TenantId) || string.IsNullOrWhiteSpace(connection.ClientId) || string.IsNullOrWhiteSpace(connection.ClientSecret))
                    throw new ArgumentNullException(@"One or more required connection values missing:", $"{nameof(connection.TenantId)}, {nameof(connection.ClientId)}, {nameof(connection.ClientSecret)}.");
                break;
            case AuthenticationProviders.UsernamePassword:
                if (string.IsNullOrWhiteSpace(connection.Username) || string.IsNullOrWhiteSpace(connection.Password) || string.IsNullOrWhiteSpace(connection.TenantId) || string.IsNullOrWhiteSpace(connection.ClientId))
                    throw new ArgumentNullException(@"One or more required connection values missing:", $"{nameof(connection.Username)}, {nameof(connection.Password)}, {nameof(connection.TenantId)}, {nameof(connection.ClientId)}.");
                break;
        }

        if (string.IsNullOrWhiteSpace(input.To))
            throw new ArgumentNullException(@"One or more required message values missing:", $"{nameof(input.To)}");
    }

    [ExcludeFromCodeCoverage(Justification = "Can't get cert and clientsecret.")]
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

    private static List<Recipient> GetRecipients(string to)
    {
        return to.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(receiver => new Recipient { EmailAddress = new EmailAddress { Address = receiver.Replace(" ", "") } }).ToList();
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

    private static void CleanUpTempFiles()
    {
        try
        {
            foreach (var file in tempFilePaths)
                if (File.Exists(file))
                    File.Delete(file);
        }
        catch (Exception)
        {
            // Do nothing if e.g. file cannot be deleted.
        }
    }

    private static async Task<Message> GetAttachments(string from, Attachments[] attachments, Options options, GraphServiceClient client, Message message, CancellationToken cancellationToken)
    {
        var attachmentList = new List<Attachment>();
        var fileList = new List<string>();
        var containLargeFile = false;

        foreach (var attachment in attachments)
        {
            switch (attachment.AttachmentType)
            {
                case AttachmentTypes.FileAttachment:
                    // If the path ends in a directory, all files in that folder with given attachment.FileMask are added as attachments.
                    string[] files = null;
                    if (File.Exists(attachment.FilePath))
                        files = new[] { attachment.FilePath };
                    else if (Directory.Exists(attachment.FilePath) && !File.Exists(attachment.FilePath))
                        files = Directory.GetFiles(attachment.FilePath, attachment.FileMask);

                    if (files != null && files.Length > 0)
                    {
                        foreach (var file in files)
                        {
                            if (new FileInfo(file).Length > 3 * 1024 * 1024)
                                containLargeFile = true;

                            fileList.Add(file);
                        }
                    }
                    else
                        if (options.ThrowExceptionIfAttachmentNotFound)
                        throw new Exception($"No files found in directory {attachment.FilePath}.");

                    break;
                case AttachmentTypes.AttachmentFromString:
                    var tempFilePath = CreateTemporaryFile(attachment);
                    if (new FileInfo(tempFilePath).Length > 3 * 1024 * 1024)
                        containLargeFile = true;

                    fileList.Add(tempFilePath);
                    tempFilePaths.Add(tempFilePath);
                    break;
            }
        }

        // Upload (large) or prepare attachment (small)
        if (containLargeFile)
        {
            //Create draft message
            message = string.IsNullOrWhiteSpace(from)
                        ? await client.Me.Messages.PostAsync(message, cancellationToken: cancellationToken)
                        : await client.Users[from].Messages.PostAsync(message, cancellationToken: cancellationToken);

            foreach (var file in fileList)
            {
                var maxSliceSize = 320 * 1024;
                var fileName = Path.GetFileName(file);
                UploadSession uploadSession;
                using var fileStream = File.OpenRead(file);

                if (string.IsNullOrEmpty(from))
                {
                    var uploadRequestBody = new Microsoft.Graph.Me.Messages.Item.Attachments.CreateUploadSession.CreateUploadSessionPostRequestBody
                    {
                        AttachmentItem = new AttachmentItem
                        {
                            AttachmentType = AttachmentType.File,
                            Name = fileName,
                            Size = fileStream.Length,
                            ContentType = "application/octet-stream"
                        },
                    };
                    uploadSession = await client.Me.Messages[message.Id].Attachments.CreateUploadSession.PostAsync(uploadRequestBody, cancellationToken: cancellationToken);
                }
                else
                {
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
                    uploadSession = await client.Users[from].Messages[message.Id].Attachments.CreateUploadSession.PostAsync(uploadRequestBody, cancellationToken: cancellationToken);
                }

                var fileUploadTask = new LargeFileUploadTask<FileAttachment>(uploadSession, fileStream, maxSliceSize);
                var totalLength = fileStream.Length;
                var largeAttachmentUploadResult = await fileUploadTask.UploadAsync(cancellationToken: cancellationToken);

                if (!largeAttachmentUploadResult.UploadSucceeded)
                    throw new Exception($@"Failed to upload large attachment ""{file}"".");
            }
        }
        else
        {
            foreach (var file in fileList)
            {
                var fileName = Path.GetFileName(file);
                var littleStream = File.ReadAllBytes(file);
                attachmentList.Add(new Attachment
                {
                    OdataType = "#microsoft.graph.fileAttachment",
                    Name = fileName,
                    AdditionalData = new Dictionary<string, object> { { "contentBytes", Convert.ToBase64String(littleStream) } },
                });
            }
            message.Attachments = attachmentList;
        }

        return message;
    }

    private static async Task<Result> SendExchangeEmail(Input input, Connection connection, Options options, CancellationToken cancellationToken)
    {
        try
        {
            using var client = CreateGraphServiceClient(connection);
            var message = new Message
            {
                Subject = input.Subject,
                Body = new ItemBody
                {
                    ContentType = input.IsMessageHtml ? BodyType.Html : BodyType.Text,
                    Content = input.Message,
                },
                ToRecipients = string.IsNullOrWhiteSpace(input.To) ? new() : GetRecipients(input.To),
                CcRecipients = string.IsNullOrWhiteSpace(input.Cc) ? new() : GetRecipients(input.Cc),
                BccRecipients = string.IsNullOrWhiteSpace(input.Bcc) ? new() : GetRecipients(input.Bcc),
                Importance = GetImportance(input.Importance),
            };

            if (input.Attachments != null && input.Attachments.Length > 0)
                message = await GetAttachments(input.From, input.Attachments, options, client, message, cancellationToken);

            if (string.IsNullOrWhiteSpace(input.From))
            {
                if (message.Id is null)
                {
                    var requestBody = new SendMailPostRequestBody() { Message = message, SaveToSentItems = input.SaveToSentItems };
                    await client.Me.SendMail.PostAsync(requestBody, cancellationToken: cancellationToken);
                }
                else
                    await client.Me.Messages[message.Id].Send.PostAsync(cancellationToken: cancellationToken);
            }
            else
            {
                if (message.Id is null)
                {
                    var userRequestBody = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody() { Message = message, SaveToSentItems = input.SaveToSentItems };
                    await client.Users[input.From].SendMail.PostAsync(userRequestBody, cancellationToken: cancellationToken);
                }
                else
                    await client.Users[input.From].Messages[message.Id].Send.PostAsync(cancellationToken: cancellationToken);
            }

            CleanUpTempFiles();

            return new Result(true, $"Email sent successfully.");
        }
        catch (Exception ex)
        {
            if (options.ThrowExceptionOnFailure)
                throw;

            return new Result(false, $"Failed to send an email. {ex.Message}");
        }
    }
}