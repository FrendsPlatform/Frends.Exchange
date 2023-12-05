using Azure.Identity;
using Frends.Exchange.ReadEmail.Definitions;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Frends.Exchange.ReadEmail;

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
    /// [Documentation](https://tasks.frends.com/tasks/frends-tasks/Frends.Exchange.ReadEmail)
    /// </summary>
    /// <param name="connection">Parameters for establishing a connection.</param>
    /// <param name="input">Email content</param>
    /// <param name="options">Options for controlling the behavior of this Task.</param>
    /// <param name="cancellationToken">Token received from Frends to cancel this Task.</param>
    /// <returns>Object { bool Success, string Data }</returns>
    public static async Task<Result> ReadEmail([PropertyTab] Connection connection, [PropertyTab] Input input, [PropertyTab] Options options, CancellationToken cancellationToken)
    {
        var resultList = new List<ResultObject>();
        var errors = new List<dynamic>();

        try
        {
            InputCheck(connection, input);
            GraphServiceClient client = CreateGraphServiceClient(connection);
            var messageCollectionResponse = await GetMessageCollectionResponse(input, client, cancellationToken);

            if (messageCollectionResponse.Value != null)
            {
                foreach (var message in messageCollectionResponse.Value)
                {
                    var resultObject = new ResultObject()
                    {
                        Id = message.Id,
                        ParentFolderId = message.ParentFolderId,
                        From = message.From?.EmailAddress?.Address,
                        Sender = message.Sender?.EmailAddress?.Address,
                        ToRecipients = message.ToRecipients?.Where(r => r != null && r.EmailAddress != null).Select(r => r.EmailAddress.Address).ToList(),
                        CcRecipients = message.CcRecipients?.Where(r => r != null && r.EmailAddress != null).Select(r => r.EmailAddress.Address).ToList(),
                        BccRecipients = message.BccRecipients?.Where(r => r != null && r.EmailAddress != null).Select(r => r.EmailAddress.Address).ToList(),
                        ReplyTo = message.ReplyTo?.Where(r => r != null && r.EmailAddress != null).Select(r => r.EmailAddress.Address).ToList(),
                        Subject = message.Subject,
                        ContentType = message.Body?.ContentType.ToString(),
                        Content = message.Body?.Content,
                        Categories = message.Categories,
                        Importance = message.Importance.ToString(),
                        IsDraft = message.IsDraft ?? false,
                        IsRead = message.IsRead ?? false,
                        HasAttachments = message.HasAttachments ?? false,
                        Extensions = message.Extensions?.Where(e => e != null).Select(e => e.Id).ToList()
                    };

                    if (input.DownloadAttachments && message.HasAttachments is true)
                        resultObject.Attachments = await DownloadAttachments(input.FileExistHandler, input.DestinationDirectory, input.CreateDirectory, message, client, cancellationToken);

                    resultList.Add(resultObject);

                    // Email won't be marked as read without doing it manually
                    if (input.UpdateReadStatus)
                        await UpdateMessageRead(input.From, message.Id, client, cancellationToken);
                }
            }
        }
        catch (Exception ex)
        {
            if (options.ThrowExceptionOnFailure)
                throw;
            else
                errors.Add(ex);
        }

        return new Result(errors.Count <= 0, resultList, errors);
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

        if (input.DownloadAttachments && string.IsNullOrWhiteSpace(input.DestinationDirectory))
            throw new ArgumentNullException($@"{nameof(input.DownloadAttachments)} is set to true, but {nameof(input.DestinationDirectory)} is missing.");

        if (!string.IsNullOrWhiteSpace(input.DestinationDirectory) && !input.CreateDirectory && !Directory.Exists(input.DestinationDirectory))
            throw new Exception($@"{nameof(input.DestinationDirectory)} is set, but the directory {input.DestinationDirectory} does not exist. Set {nameof(input.CreateDirectory)} to true to create the specified directory.");
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

    private static async Task<MessageCollectionResponse> GetMessageCollectionResponse(Input input, GraphServiceClient client, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(input.From))
        {
            return await client.Me.Messages.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Count = true;
                requestConfiguration.QueryParameters.Filter = string.IsNullOrWhiteSpace(input.Filter) ? null : input.Filter;
                requestConfiguration.QueryParameters.Skip = input.Skip;
                requestConfiguration.QueryParameters.Top = input.Top > 0 ? input.Top : null;
                requestConfiguration.QueryParameters.Orderby = string.IsNullOrWhiteSpace(input.Orderby) ? null : input.Orderby.Split(new[] { "\", \"" }, StringSplitOptions.None);
                requestConfiguration.QueryParameters.Expand = string.IsNullOrWhiteSpace(input.Expand) ? null : input.Expand.Split(new[] { "\", \"" }, StringSplitOptions.None);
                if (input.Headers != null && input.Headers.Length > 0)
                    foreach (var header in input.Headers)
                        requestConfiguration.Headers.Add(header.HeaderName, header.HeaderValues.ToArray());
            }, cancellationToken);
        }

        return await client.Users[input.From].Messages.GetAsync((requestConfiguration) =>
        {
            requestConfiguration.QueryParameters.Count = true;
            requestConfiguration.QueryParameters.Filter = string.IsNullOrWhiteSpace(input.Filter) ? null : input.Filter;
            requestConfiguration.QueryParameters.Skip = input.Skip;
            requestConfiguration.QueryParameters.Top = input.Top > 0 ? input.Top : null;
            requestConfiguration.QueryParameters.Orderby = string.IsNullOrWhiteSpace(input.Orderby) ? null : input.Orderby.Split(new[] { "\", \"" }, StringSplitOptions.None);
            requestConfiguration.QueryParameters.Expand = string.IsNullOrWhiteSpace(input.Expand) ? null : input.Expand.Split(new[] { "\", \"" }, StringSplitOptions.None);
            if (input.Headers != null && input.Headers.Length > 0)
                foreach (var header in input.Headers)
                    requestConfiguration.Headers.Add(header.HeaderName, header.HeaderValues.ToArray());
        }, cancellationToken);
    }

    private static async Task<List<Attachments>> DownloadAttachments(FileExistHandlers fileExistHandler, string filePath, bool createDir, Message message, GraphServiceClient client, CancellationToken cancellationToken)
    {
        if (createDir && !Directory.Exists(filePath))
            Directory.CreateDirectory(filePath);

        var attachmentsList = new List<Attachments>();
        var attachments = client.Me.Messages[message.Id].Attachments.GetAsync((requestConfiguration) =>
        {
            requestConfiguration.QueryParameters.Expand = new string[] { "microsoft.graph.itemattachment/item" };
        }, cancellationToken: cancellationToken).Result;

        foreach (var attachment in attachments.Value)
        {
            if (attachment is FileAttachment fileAttachment)
            {
                var createFileFromBytes = await CreateFileFromBytes(fileExistHandler, fileAttachment.ContentBytes, Path.Combine(filePath, fileAttachment.Name), cancellationToken);
                attachmentsList.Add(new Attachments()
                {
                    Id = fileAttachment.Id,
                    FilePath = createFileFromBytes,
                    Size = fileAttachment.Size,
                    OdataType = fileAttachment.OdataType,
                    Content = null,
                });
            }
            // ItemAttachment represents an email with its own attachments, attached to another email. This will downloads the attachments of the attached email.
            if (attachment is ItemAttachment itemAttachment)
            {
                if (itemAttachment.Item is Message item)
                    foreach (var innerAttachment in item.Attachments)
                    {
                        var itemAsFileAttachment = innerAttachment as FileAttachment;
                        var itemBytes = itemAsFileAttachment.ContentBytes;
                        var createFileFromBytes = await CreateFileFromBytes(fileExistHandler, itemBytes, Path.Combine(filePath, innerAttachment.Name), cancellationToken);

                        attachmentsList.Add(new Attachments()
                        {
                            Id = itemAttachment.Id,
                            FilePath = createFileFromBytes,
                            Size = itemAttachment.Size,
                            OdataType = itemAttachment.OdataType,
                            Content = null,
                        });
                    }
            }
        }
        return attachmentsList;
    }

    private static async Task<string> CreateFileFromBytes(FileExistHandlers fileExistHandler, byte[] contentBytes, string filePath, CancellationToken cancellationToken)
    {
        if (!File.Exists(filePath))
        {
            using (var fs = new FileStream(filePath, FileMode.CreateNew))
                await fs.WriteAsync(contentBytes, cancellationToken);
            return filePath;
        }
        else
        {
            switch (fileExistHandler)
            {
                case FileExistHandlers.Skip:
                    return $@"The file {filePath} already exists. Download skipped.";
                case FileExistHandlers.Rename:
                    filePath = GetUniqueFilePath(filePath);
                    using (var fs = new FileStream(filePath, FileMode.CreateNew))
                        await fs.WriteAsync(contentBytes, cancellationToken);
                    return filePath;
                case FileExistHandlers.Append:
                    using (var fs = new FileStream(filePath, FileMode.Append))
                        await fs.WriteAsync(contentBytes, cancellationToken);
                    return filePath;
                case FileExistHandlers.OverWrite:
                    using (var fs = new FileStream(filePath, FileMode.Create))
                        await fs.WriteAsync(contentBytes, cancellationToken);
                    return filePath;
                default: throw new Exception("An exception occurred while trying to handle an already existing file.");
            }
        }
    }

    private static string GetUniqueFilePath(string filePath)
    {
        if (File.Exists(filePath))
        {
            var directory = Path.GetDirectoryName(filePath);
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var extension = Path.GetExtension(filePath);
            var count = 1;

            while (File.Exists(Path.Combine(directory, $"{fileName}({count}){extension}")))
                count++;

            return Path.Combine(directory, $"{fileName}({count}){extension}");
        }

        return filePath;
    }

    private static async Task UpdateMessageRead(string from, string messageId, GraphServiceClient client, CancellationToken cancellationToken)
    {
        var requestBody = new Message { IsRead = true };

        if (string.IsNullOrWhiteSpace(from))
            await client.Me.Messages[messageId].PatchAsync(requestBody, cancellationToken: cancellationToken);
        else
            await client.Users[from].Messages[messageId].PatchAsync(requestBody, cancellationToken: cancellationToken);
    }
}