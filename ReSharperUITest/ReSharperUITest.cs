using System;
using System.Linq;
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

        // todo: add start Visual Studio before all TestCases
        // todo: add close Visual Studio after all TestCases

        [SetUp]
        public void BeforeEachTest()
        {
            // Create a session for Desktop
            var desktopCapabilities = new DesiredCapabilities();
            desktopCapabilities.SetCapability("app", "Root");
            _desktopSession =
                new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), desktopCapabilities);
            Assert.IsNotNull(_desktopSession);
            // find Microsoft Visual Studio window todo replace with partially Element Name
            var visualStudioWindow = _desktopSession.FindElementByName("ReSharperUITest - Microsoft Visual Studio ");
            var visualStudioTopLevelWindowHandle =
                (int.Parse(visualStudioWindow.GetAttribute("NativeWindowHandle"))).ToString("x");

            // Create session for already running Microsoft Visual Studio
            var appCapabilities = new DesiredCapabilities();
            appCapabilities.SetCapability("appTopLevelWindow", visualStudioTopLevelWindowHandle);
            _driver = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
            Assert.IsNotNull(_driver);

            // Set managed memory usage to not displayed this means that test will always starts on the same point
            if (IsManagedMemoryDisplayed())
            {
                ToggleManagedMemoryCheckBox();
            }

        }

        [TearDown]
        public void AfterEachTest()
        {
            //_driver?.Quit();
            //_driver = null;

            // close "Options" window if it is not closed
            var optionWindow = _driver.FindElementsByName("Options").FirstOrDefault();
            optionWindow?.FindElementsByName("Cancel").FirstOrDefault()?.Click();

        }

        [Test]
        public void OpenReSharperOptions()
        {
            if (IsManagedMemoryDisplayed())
            {
                ToggleManagedMemoryCheckBox();
            }
        }

        /// <summary>
        /// Open ReSharper Option and toggle "Show managed memory usage in status bar" checkbox
        /// </summary>
        private void ToggleManagedMemoryCheckBox()
        {
            // Click on ReSharper menu 
            _driver.FindElementByName("ReSharper").Click();
            // Open ReSharper Option Dialog
            //_driver.FindElementByXPath($"//MenuItem[starts-with(@Name, \"Options…\")]").Click();
            // for unknown reason FindElementByXPath($"//MenuItem[starts-with(@Name, \"Options…\")]") not able to find "Options…" menu item
            var reSharperMenuItems = _driver.FindElementsByClassName("MenuItem");
            reSharperMenuItems.First(p => p.Text == "Options…").Click();
            // Find Option Window
            var optionWindow = _driver.FindElementByName("Options");
            Assert.IsNotNull(optionWindow);

            // EMK it is not necessary to click on "Settings" but 
            //   task description contain the following:
            // "Важно: закладка Genral скорее всего будет выбрана по дефолту, но тест должен уметь
            //  находить опцию, если открыта любая закладка"
            // So switch to "Settings" to emulate the situation that some other tab is opened
            var settings = optionWindow.FindElementByName("Settings");
            Assert.IsNotNull(settings);
            settings.Click();

            // Locate "General" tab
            var general = optionWindow.FindElementByName("General");
            Assert.IsNotNull(general);
            general.Click();

            // Locate "Show managed _memory usage in status bar
            var managedMemoryCheckBox = optionWindow.FindElementByName("Show managed _memory usage in status bar");
            Assert.IsNotNull(managedMemoryCheckBox);
            // toggle "Show managed _memory usage in status bar"
            managedMemoryCheckBox.Click();

            // Save & Close "Option" ReSharper Option window
            var saveButton = optionWindow.FindElementByName("Save");
            Assert.IsNotNull(saveButton);
            saveButton.Click();
        }

        private bool IsManagedMemoryDisplayed()
        {
            return !bool.Parse(_driver.FindElementByClassName("RichTextPresenter").GetAttribute("IsOffscreen"));
        }
    }
}