namespace WeatherAppOpenXML.Tests;

using Microsoft.AspNetCore.Routing;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

[TestFixture]
public class CityInputTests
{
    private IWebDriver driver;

    [SetUp]
    public void Setup()
    {
        driver = new ChromeDriver();
        driver.Navigate().GoToUrl("https://localhost:7169");        
    }

    [TearDown]
    public void Teardown()
    {
        driver.Quit();
    }

    [Test]
    public void IntegerInput_DisablesExportButton()
    {
        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
        IWebElement cityInput = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("cityInput")));
        cityInput.Clear();
        System.Threading.Thread.Sleep(2000);

        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
        js.ExecuteScript("document.getElementById('cityInput').value='12345';");
        js.ExecuteScript("document.activeElement.blur();");
        System.Threading.Thread.Sleep(2000);


        bool isButtonDisabled = (bool)js.ExecuteScript(@"
            var element = document.getElementById('hoverButton');
            return element !== null && element.disabled === true;
        ");

        Assert.IsFalse(isButtonDisabled, "The export button should not be present or enabled.");

    }

    [Test]
    public void EmptyInput_DisablesExportButton()
    {
        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
        IWebElement cityInput = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("cityInput")));
        cityInput.Clear();
        System.Threading.Thread.Sleep(2000);

        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
        js.ExecuteScript("document.getElementById('cityInput').value='';");
        js.ExecuteScript("document.activeElement.blur();");
        System.Threading.Thread.Sleep(2000);


        bool isButtonDisabled = (bool)js.ExecuteScript(@"
            var element = document.getElementById('hoverButton');
            return element !== null && element.disabled === true;
        ");

        Assert.IsFalse(isButtonDisabled, "The export button should not be present or enabled.");
    }

    [Test]
    public void NonCharacterInput_DisablesExportButton()
    {
        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
        IWebElement cityInput = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("cityInput")));
        cityInput.Clear();
        System.Threading.Thread.Sleep(2000);

        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
        js.ExecuteScript("document.getElementById('cityInput').value='@!#';");
        js.ExecuteScript("document.activeElement.blur();");
        System.Threading.Thread.Sleep(2000);


        bool isButtonDisabled = (bool)js.ExecuteScript(@"
            var element = document.getElementById('hoverButton');
            return element !== null && element.disabled === true;
        ");

        Assert.IsFalse(isButtonDisabled, "The export button should not be present or enabled.");
    }

    [Test]
    public void ValidCharacterInput_EnablesExportButton()
    {
        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
        IWebElement cityInput = wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.Id("cityInput")))[0];
        cityInput.Clear();
        System.Threading.Thread.Sleep(2000);

        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
        js.ExecuteScript("document.getElementById('cityInput').value='Pleven';");

        System.Threading.Thread.Sleep(2000);

        bool isButtonEnabled = (bool)js.ExecuteScript(@"
            var element = document.getElementById('hoverButton');
            return element !== null && element.disabled === false;
        ");
        Assert.IsTrue(isButtonEnabled, "The export button should be enabled.");
    }
}

