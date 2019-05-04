using System;
using NUnit.Framework;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;

namespace ReSharperUITest
{
    [TestFixture]
    public class ReSharperTest
    {
        private const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";

        private WindowsDriver<WindowsElement> _driver;
        private WindowsDriver<WindowsElement> _desktopSession;

        [SetUp]
        public void TestInit()
        {
            // Create a session for Desktop
            var desktopCapabilities = new DesiredCapabilities();
            desktopCapabilities.SetCapability("app", "Root");
            _desktopSession = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), desktopCapabilities);
            Assert.IsNotNull(_desktopSession);
            // find Microsoft Visual Studio window todo replace with partially Element Name
            var visualStudioWindow = _desktopSession.FindElementByName("ReSharperUITest - Microsoft Visual Studio ");
            var visualStudioTopLevelWindowHandle = (int.Parse(visualStudioWindow.GetAttribute("NativeWindowHandle"))).ToString("x");

            // Create session for already running Microsoft Visual Studio
            var appCapabilities = new DesiredCapabilities();
            appCapabilities.SetCapability("appTopLevelWindow", visualStudioTopLevelWindowHandle);
            _driver = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
            Assert.IsNotNull(_driver);
        }

        [TearDown]
        public void TestCleanup()
        {
            if (_driver != null)
            {
                _driver.Quit();
                _driver = null;
            }
        }

        [Test]
        public void OpenResharperOptions()
        {
            
            _driver.FindElementByName("ReSharper").Click(); // open ReSharper menu
            _driver.FindElementByName("Options...").Click(); // open ReSharper menu
            //// LegacyIAccessible.Name: "General"
            //_driver.FindElementByName("General").Click();
            //// LegacyIAccessible.Name: "Show managed _memory usage in status bar"
            ////  Probably it is necessary to verify the status and check/uncheck
            //_driver.FindElementByName("Show managed _memory usage in status bar").Click();
        }
    }
}
