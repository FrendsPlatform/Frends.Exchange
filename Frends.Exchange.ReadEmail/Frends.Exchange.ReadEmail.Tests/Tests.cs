using NUnit.Framework;
using Frends.Exchange.ReadEmail;
using System;
using Microsoft.Exchange.WebServices.Data;
using Frends.Exchange.ReadEmail.Definitions;
using System.Net.Security;
using System.Reflection;
using Moq;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;

namespace Frends.Exchange.ReadEmail.Tests
{
    public class Tests
    {
        private const string serverAddressInCorrectFormat = "https://servername/ews/exchange.amsx";
        private const string serverAddressAsFaulty = "ksdfjdosifsfsdsc/exchange.asddd";
        private const string expectedDummyTextFileContent = "This is a dummy file.";


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
        public void ConnectToExchangeService()
        {
            TestSettingExchangeServerVersion(ExchangeServerVersion.Exchange2007_SP1, ExchangeVersion.Exchange2007_SP1);
            TestSettingExchangeServerVersion(ExchangeServerVersion.Exchange2010, ExchangeVersion.Exchange2010);
            TestSettingExchangeServerVersion(ExchangeServerVersion.Exchange2010_SP1, ExchangeVersion.Exchange2010_SP1);
            TestSettingExchangeServerVersion(ExchangeServerVersion.Exchange2010_SP2, ExchangeVersion.Exchange2010_SP2);
            TestSettingExchangeServerVersion(ExchangeServerVersion.Exchange2013_SP1, ExchangeVersion.Exchange2013_SP1);
            TestSettingExchangeServerVersion(ExchangeServerVersion.Office365, ExchangeVersion.Exchange2013_SP1);
        }

        /// <summary>
        /// 
        /// </summary>
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

        [Test]
        public void ExchangeCertificateValidationCallbackIsSetCorrectly()
        {

        }
    }
}