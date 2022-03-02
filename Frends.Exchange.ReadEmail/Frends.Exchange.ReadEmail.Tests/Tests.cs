using NUnit.Framework;
using System;
using Microsoft.Exchange.WebServices.Data;
using Frends.Exchange.ReadEmail.Definitions;
using System.Net.Security;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading;
using Newtonsoft.Json;
using System.Text;

namespace Frends.Exchange.ReadEmail.Tests
{
    public class Tests
    {
        private const string serverAddressInCorrectFormat = "https://servername/ews/exchange.amsx";
        private const string serverAddressAsFaulty = "ksdfjdosifsfsdsc/exchange.asddd";
        private const string expectedDummyTextFileContent = "This is a dummy file.";
        private const string settingsEnvVarName = "EXCHANGE_SETTINGS_FOR_TESTING";

        public static ExchangeSettings GetSettingsFromEnvironment()
        {
            var base64str = Environment.GetEnvironmentVariable(settingsEnvVarName);
            if (string.IsNullOrEmpty(base64str))
            {
                throw new InvalidOperationException(
                    $"The environment variable \"{settingsEnvVarName}\" appears to be empty. This could be because it either is really empty, or the variable doesn't exist.");
            }

            string json = "";
            try
            {
                json = Encoding.UTF8.GetString(Convert.FromBase64String(base64str));
            }
            catch (FormatException formatEx)
            {
                throw new Exception($"Couldn't decode environment variable \"{settingsEnvVarName}\" from base64. It might be because it isn't base64-encoded properly.", formatEx);
            }
            catch(Exception ex)
            {
                throw new Exception($"Couldn't decode environment variable \"{settingsEnvVarName}\" from base64.", ex);
            }

            try
            {
                return JsonConvert.DeserializeObject<ExchangeSettings>(json);
            }
            catch (Exception ex)
            {
                throw new Exception($"Couldn't deserialize the environment variable \"{settingsEnvVarName}\". Ensure it's in correct format.", ex);
            }
        }

        [Test]
        public void ConnectToExchangeServiceRejectsFaultyServerAddress()
        {
            var settings = new ExchangeSettings()
            {
                ServerAddress = serverAddressAsFaulty
            };
            
            Assert.Throws(typeof(UriFormatException), () => 
            { 
                Util.ConnectToExchangeService(settings); 
            });
        }

        [Test]
        public void ConnectToExchangeServiceMethodSetsCorrectExchangeVersion()
        {
            TestSettingExchangeServerVersion(ExchangeServerVersion.Exchange2007_SP1, ExchangeVersion.Exchange2007_SP1);
            TestSettingExchangeServerVersion(ExchangeServerVersion.Exchange2010, ExchangeVersion.Exchange2010);
            TestSettingExchangeServerVersion(ExchangeServerVersion.Exchange2010_SP1, ExchangeVersion.Exchange2010_SP1);
            TestSettingExchangeServerVersion(ExchangeServerVersion.Exchange2010_SP2, ExchangeVersion.Exchange2010_SP2);
            TestSettingExchangeServerVersion(ExchangeServerVersion.Exchange2013_SP1, ExchangeVersion.Exchange2013_SP1);
            TestSettingExchangeServerVersion(ExchangeServerVersion.Office365, ExchangeVersion.Exchange2013_SP1);
        }

        [Test]
        public void ConnectToExchangeServiceSetsExpectedValidationCallbackMethod()
        {
            Util.ConnectToExchangeService(new ExchangeSettings()
            {
                ServerAddress = serverAddressInCorrectFormat,
            });

            var method = System.Net.ServicePointManager.ServerCertificateValidationCallback;
            Assert.IsNotNull(method?.Method, "ConnectToExchangeService method didn't set ServicePointManager.ServerCertificateValidationCallback even though it should.");
            Assert.IsTrue(method!.Invoke(this, null, null, SslPolicyErrors.None));
            Assert.IsTrue(method!.Invoke(this, null, null, SslPolicyErrors.RemoteCertificateChainErrors));
            Assert.IsFalse(method!.Invoke(this, null, null, SslPolicyErrors.RemoteCertificateNameMismatch));
        }

        /// <summary>
        /// Method for testing that the attachments get saved to filesystem as intended, 
        /// with their content as intended.
        /// </summary>
        [Test]
        public void AttachmentsGetSavedAsShould()
        {
            var service = new ExchangeService();
            var tempDirPath = Path.Combine(Path.GetTempPath(), "ReadEmailTests");
            var msgAndFiles = CreateEmailMessageWithAttachments(tempDirPath, service);

            var dummyDirPath = Path.Combine(tempDirPath, "DummySaveDir");
            Directory.CreateDirectory(dummyDirPath);
            var options = new ExchangeOptions()
            {
                AttachmentSaveDirectory = dummyDirPath,
            };
            Util.SaveAttachments(msgAndFiles.Item1.Attachments, options);

            foreach (var fileInfo in msgAndFiles.Item2)
            {
                Assert.IsTrue(fileInfo.Exists);
                var content = File.ReadAllText(fileInfo.FullName);
                Assert.AreEqual(expectedDummyTextFileContent, content);
            }

            Directory.Delete(tempDirPath, true);
        }

        private static (EmailMessage, IEnumerable<FileInfo>) CreateEmailMessageWithAttachments(string tempDirPath, ExchangeService service)
        {
            var files = CreateTempTextFiles(tempDirPath, 5);

            var msg = new EmailMessage(service);
            foreach (var file in files)
            {
                msg.Attachments.AddFileAttachment(file.FullName);
            }

            return (msg, files);
        }

        private static IEnumerable<FileInfo> CreateTempTextFiles(string dirPath, int count)
        {
            if (count <= 0)
            {
                throw new ArgumentException("Count must be over 0.");
            }

            int counter = 0;
            while (counter < count)
            {
                yield return CreateTempTextFile(dirPath);
                counter++;
            }
        }

        private static FileInfo CreateTempTextFile(string dirPath)
        {
            Directory.CreateDirectory(dirPath);
            var info = new FileInfo(Path.Combine(dirPath, Guid.NewGuid().ToString() + ".txt"));
            using var writer = File.CreateText(info.FullName);
            writer.Write(expectedDummyTextFileContent);
            writer.Flush();
            return info;
        }

        /// <summary>
        /// Initiates a new ExchangeService object with 
        /// <see cref="Frends.Exchange.ReadEmail.Util.ConnectToExchangeService"/> method
        /// and tests if the Exchange server version in that object is what is expected.
        /// If no <paramref name="expectedServerVersion"/> is set, this method asserts
        /// that the resulted version equals <paramref name="setServerVersion"/> value.
        /// </summary>
        private static void TestSettingExchangeServerVersion(ExchangeServerVersion setServerVersion, ExchangeVersion expectedVersion)
        {
            var service = CreateCertainTypedServiceWithConnectMethod(setServerVersion);

            Assert.AreEqual(expectedVersion, service.RequestedServerVersion,
                $"Expected service.RequestedServerVersion property value to be {expectedVersion}, but it was {service.RequestedServerVersion}");
        }

        private static ExchangeService CreateCertainTypedServiceWithConnectMethod(ExchangeServerVersion exchangeVersion)
        {
            return Util.ConnectToExchangeService(new ExchangeSettings()
            {
                ExchangeServerVersion = exchangeVersion,
                ServerAddress = serverAddressInCorrectFormat
            });
        }

        /// <summary>
        /// Tests that Util
        /// <list type="number">
        /// <item>contains a method with a name RedirectionUrlValidationCallback</item>
        /// <item>the method returns a <see cref="bool"/></item>
        /// <item>the method takes one <see cref="string"/> param</item>
        /// <item>the method returns true when the passed param url has https scheme</item>
        /// <item>the method returns false when the passed param url doesn't have https scheme</item>
        /// </list>
        /// </summary>
        [Test]
        public void RedirectionUrlValidationCallbackReturnsAsExpected()
        {
            var args = new[]
            {
                "http://not.nice.url.com"
            };
            var method = typeof(Util).GetMethod("RedirectionUrlValidationCallback");
            Assert.IsNotNull(method);
            Assert.AreEqual(typeof(bool), method!.ReturnType);
            var methodParam = method.GetParameters();
            Assert.IsNotNull(methodParam);
            Assert.NotZero(methodParam.Length);
            Assert.AreEqual(typeof(string), methodParam[0].ParameterType);

            var result = method.Invoke(null, args);
            Assert.IsInstanceOf(typeof(bool), result, "The method provided for RedirectionUrlValidationCallback doesn't return a bool.");
            Assert.IsFalse((bool)result!);

            args[0] = "https://nice.url.com";
            result = method.Invoke(null, args);
            Assert.IsTrue((bool)result!);
        }

        private static EmailMessage SendTestEmail(ExchangeSettings settings)
        {
            var tempDirPath = Path.Combine(Path.GetTempPath(), "emailsendtest");
            Directory.CreateDirectory(tempDirPath);

            var service = Util.ConnectToExchangeService(settings);
            var msgAndFiles = CreateEmailMessageWithAttachments(tempDirPath, service);

            msgAndFiles.Item1.Subject = $"Test email {Guid.NewGuid()}";
            msgAndFiles.Item1.Body = "Hello there! This is a test email, for testing purposes.";
            msgAndFiles.Item1.ToRecipients.Add(settings.Username);

            var sendTask = msgAndFiles.Item1.Send();
            sendTask.Wait(5000);

            Directory.Delete(tempDirPath, true);

            return msgAndFiles.Item1;
        }

        [Test]
        public static void TestReadingMails()
        {
            var settings = GetSettingsFromEnvironment();
            var options = new ExchangeOptions()
            {
                AttachmentSaveDirectory = Path.Combine(Path.GetTempPath(), "reademailtest"),
                MaxEmails = 500
            };
            Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "reademailtest"));

            var sentEmail = SendTestEmail(settings);
            
            Thread.Sleep(5000);

            var readTask = Exchange.ReadEmail(settings, options);
            readTask.Wait(10000);
            var result = readTask.Result;

            EmailMessageResult? receivedEmail = result.Where(msg => msg.Subject == sentEmail.Subject).FirstOrDefault();

            Assert.IsNotNull(receivedEmail);

            RemoveEmailMessage(Util.ConnectToExchangeService(settings), receivedEmail!.Id);
        }

        public static void RemoveEmailMessage(ExchangeService service, string msgId)
        {
            var response = service.DeleteItems(new[] { new ItemId(msgId) }, DeleteMode.HardDelete, null, null);
        }
    }
}