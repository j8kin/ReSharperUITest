using System;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.Text.RegularExpressions;
using System.Threading;
using NUnit.Framework;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;

namespace ReSharperUITest
{
    [TestFixture]
    public class ReSharperTest
    {
        private const string WinAppDriverPath =
            "\"c:\\Program Files (x86)\\Windows Application Driver\\WinAppDriver.exe\"";

        private const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";

        private WindowsDriver<WindowsElement> _driver;
        private Process _winAppDriverProcess;

        [OneTimeSetUp]
        public void BeforeAllTests()
        {
            // Start WinAppDriver with Administrative rights
            // User "nunittest" with password "Passw0rd!" should be created on remote host
            _winAppDriverProcess = new Process
            {
                StartInfo =
                {
                    FileName = WinAppDriverPath,
                    UseShellExecute = true,
                    Verb = "runas",
                    UserName = "nunittest"
                }
            };
            var secure = new SecureString();
            foreach (var c in "Passw0rd!".ToCharArray())
            {
                secure.AppendChar(c);
            }

            _winAppDriverProcess.StartInfo.Password = secure;
            _winAppDriverProcess.Start(); // this stuff is not verified :(

            // Start visual studio on remote Host
            // path to folder with devenv.exe must be placed into PATH
            // environment variable
            var appCapabilities = new DesiredCapabilities();
            //appCapabilities.SetCapability("appTopLevelWindow", visualStudioTopLevelWindowHandle);
            appCapabilities.SetCapability("app", "devenv.exe");
            _driver = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
            Assert.IsNotNull(_driver);

            // Set implicit timeout to 1.5 seconds to make element search to retry every 500 ms for at most three times
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1.5);

            // Pause due to my PC is too slow and it is necessary to perform additional pause to start Visual Studio
            Thread.Sleep(20000);
        }

        [OneTimeTearDown]
        public void AfterAllTests()
        {
            // Close visual studio on Remote Host
            _driver?.Quit();
            _driver = null;

            _winAppDriverProcess?.Kill(); 
        }

        [SetUp]
        public void BeforeEachTest()
        {
            // Set managed memory usage to not displayed this means that test will always starts on the same point
            if (IsManagedMemoryDisplayed())
            {
                ToggleManagedMemoryCheckBox();
            }
        }

        [TearDown]
        public void AfterEachTest()
        {
            // close "Options" window if it is not closed
            var optionWindow = _driver?.FindElementsByName("Options").FirstOrDefault();
            optionWindow?.FindElementsByName("Cancel").FirstOrDefault()?.Click();
        }

        [Test]
        public void ClickOnOptionMenu()
        {
            _driver.FindElementByName("ReSharper").Click();
            // Open ReSharper Option Dialog
            // for unknown reason FindElementByXPath($"//MenuItem[starts-with(@Text, \"Options…\")]") not able to find "Options…" menu item
            _driver.FindElementsByClassName("MenuItem").First(p=>p.Text== "Options…").Click();

        }

        [Test]
        public void SwitchBetweenTabs()
        {
            _driver.FindElementByName("ReSharper").Click();

            // Open ReSharper Option Dialog
            // for unknown reason FindElementByXPath($"//MenuItem[starts-with(@Text, \"Options…\")]") not able to find "Options…" menu item
            //  so get all children and click on child with "Options…" text
            _driver.FindElementsByClassName("MenuItem").First(p => p.Text == "Options…").Click();

            // Find Option Window
            var optionWindow = _driver.FindElementByName("Options");

            var gridElem = optionWindow.FindElementByName("Settings");
            // For unknown reason element.Click() is not working.
            // It just only mouse move to element position
            // that is why use Action for this
            new Actions(_driver).MoveToElement(gridElem).Click().Perform();

            gridElem = optionWindow.FindElementByName("General");
            // For unknown reason element.Click() is not working.
            // It just only mouse move to element position
            // that is why use Action for this
            new Actions(_driver).MoveToElement(gridElem).Click().Perform();
        }

        [Test]
        public void VerifyReSharperManagedMemoryOptionLogicAsync()
        {
            // Initially Managed Memory is not Displayed
            Assert.False(IsManagedMemoryDisplayed());

            // Toggle Managed Memory Checkbox
            ToggleManagedMemoryCheckBox();

            // Verify that Managed Memory is Displayed
            Assert.True(IsManagedMemoryDisplayed());

            // Toggle Managed Memory Checkbox
            ToggleManagedMemoryCheckBox();

            // Verify that Managed Memory is not Displayed
            Assert.False(IsManagedMemoryDisplayed());
        }

        [Test]
        public void DummyFailedTest()
        {
            // This test is necessary only for execution on remote host and 
            //   demonstrate that result is fail and provide TestResult.xml 
            //   with this fail
            Assert.True(false);
        }

        /// <summary>
        /// Open ReSharper Option and toggle "Show managed memory usage in status bar" checkbox
        /// </summary>
        private void ToggleManagedMemoryCheckBox()
        {
            // Click on ReSharper menu 
            _driver.FindElementByName("ReSharper").Click();

            // Open ReSharper Option Dialog
            //_driver.FindElementByName("ReSharper").FindElementsByXPath(".//*").First(p => p.Text == "Options…").Click();
            _driver.FindElementsByClassName("MenuItem").First(p => p.Text == "Options…").Click();

            // Find Option Window
            var optionWindow = _driver.FindElementByName("Options");

            // EMK it is not necessary to click on "Settings" but 
            //   task description contain the following:
            // "Важно: закладка Genral скорее всего будет выбрана по дефолту, но тест должен уметь
            //  находить опцию, если открыта любая закладка"
            // So switch to "Settings" to emulate the situation that some other tab is opened
            var gridElement = optionWindow.FindElementByName("Settings");
            new Actions(_driver).MoveToElement(gridElement).Click().Perform();

            // Locate "General" tab
            gridElement = optionWindow.FindElementByName("General");
            new Actions(_driver).MoveToElement(gridElement).Click().Perform();

            // Locate "Show managed _memory usage in status bar and toggle it
            optionWindow.FindElementByName("Show managed _memory usage in status bar").Click();

            // Save & Close "Option" ReSharper Option window
            optionWindow.FindElementByName("Save").Click();
        }

        /// <summary>
        /// Return status of Managed memory.
        /// 
        /// The assumption is that Managed Memory consists of any number of digit and then "MB" after space
        ///   for real test it could not work since it could be in GB, so need to confirm and update if it is true
        /// This function is for current demonstration
        /// </summary>
        /// <returns></returns>
        private bool IsManagedMemoryDisplayed()
        {
            return _driver.FindElementsByClassName("RichTextPresenter").Count != 0 && Regex.IsMatch(_driver.FindElementByClassName("RichTextPresenter").Text, "\\d+(\\.\\d+)* MB");
        }
    }
}