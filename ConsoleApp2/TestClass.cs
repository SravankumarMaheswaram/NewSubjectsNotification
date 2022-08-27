using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Guru99Demo
{
    public class TestClass
    {
        IWebDriver driver;



        [SetUp]
        public void startBrowser()
        {
            driver = new ChromeDriver("C:\\driver\\");
            System.Threading.Thread.Sleep(5000);
        }

        [Test]
        public void test()
        {
            driver.Url = "http://www.google.co.in";
            System.Threading.Thread.Sleep(5000);

            //IWebelement element = driver.FindElement(By.xpath("xpath of Webelement"));
            //Boolean status = element.Selected;

        }

        [TearDown]
        public void closeBrowser()
        {
            driver.Close();
        }

    }
}