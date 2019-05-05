using System;
using System.IO;
using System.Reflection;
using JetBrains.OsTestFramework;
using NUnit.Framework;
using System.IO.Compression;

namespace RemoteHostUITestExecution
{
    [TestFixture]
    public class RemoteHostUiTestExecution
    {
        private const string PsToolPath = @"c:\Users\Eugene\source\repos\PSTools\PsExec.exe";

        private static readonly string AssemblyDirectory =
            Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().EscapedCodeBase).AbsolutePath);

        private const string Ip = "192.168.0.84";
        private const string UserName = "nunittest";
        private const string Password = "Passw0rd!";

        private readonly string _pathToUinUnitTest = AssemblyDirectory + @"..\..\..\..\ReSharperUITest\bin\Debug";

        [OneTimeSetUp]
        public void BeforeAllTests()
        {
            // Remove previous version of ReSharperUITests.zip
            if (File.Exists(_pathToUinUnitTest + @"\ReSharperUITests.zip"))
            {
                File.Delete(_pathToUinUnitTest + @"\ReSharperUITests.zip");
            }

            // Zip all UI Tests from PathToUINUnitTest
            ZipFile.CreateFromDirectory(_pathToUinUnitTest, _pathToUinUnitTest + @"\..\ReSharperUITests.zip");
            File.Move(_pathToUinUnitTest + @"\..\ReSharperUITests.zip", _pathToUinUnitTest + @"\ReSharperUITests.zip");
        }

        [Test]
        public void ZipUiTestBinaries()
        {
            // this test is necessary only to verify that ReSharperUITests compressed successfully
            Assert.True(File.Exists(_pathToUinUnitTest + @"\ReSharperUITests.zip"));
        }

        [Test]
        public void ExecuteUiTestOnRemoteHost()
        {
            // Test Assumption C:\_work folder exist on remote Host

            using (var operatingSystem = new RemoteEnvironment(Ip, UserName, Password, PsToolPath))
            {
                // Copy nUnit test to Guest PC
                operatingSystem.CopyFileFromHostToGuest(
                    _pathToUinUnitTest + @"\ReSharperUITests.zip",
                    @"C:\_work");
                // Unzip nUnit test on Guest PC
                operatingSystem.WindowsShellInstance.DetachElevatedCommandInGuestNoRemoteOutput(
                    @"7z.exe x -aoa -oc:\_work\Output c:\_work\ReSharperUITests.zip", TimeSpan.FromSeconds(1));
                // Execute nUnit test on Guest PC
                operatingSystem.WindowsShellInstance.DetachElevatedCommandInGuestNoRemoteOutput(
                    @"nunit3-console --work=c:\_work\Output c:\_work\Output\ReSharperUITest.dll",
                    TimeSpan.FromSeconds(1));

                // Verify that nunit execution test result (TestResult.xml) exist in "c:\_work\Output"
                Assert.True(operatingSystem.FileExistsInGuest(@"c:\_work\Output\TestResult.xml"));

                // Remove previously copied test result
                if (File.Exists(_pathToUinUnitTest + @"\TestResult.xml"))
                {
                    File.Delete(_pathToUinUnitTest + @"\TestResult.xml");
                }

                // Copy Test Result from Guest to Host
                operatingSystem.CopyFileFromGuestToHost(
                    @"c:\_work\Output\TestResult.xml",
                    _pathToUinUnitTest);

                // TestContext.AddTestAttachment(_pathToUinUnitTest + @"\TestResult.xml", "NUnit Test Result on remote host");
            }
        }
    }
}