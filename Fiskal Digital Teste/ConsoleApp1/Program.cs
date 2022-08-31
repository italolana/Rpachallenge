using System;
using System.Linq;
using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;



namespace LerPlanilhaExcel
{
    class Program
    {
        
        public static string DataDir { get; private set; }

        static void Main(string[] args)
        {    //Abrir Excel
            var xls = new XLWorkbook(@"../../challenge.xlsx");
            var planilha = xls.Worksheets.First(w => w.Name == "Sheet1");
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://rpachallenge.com/");


            driver.FindElement(By.XPath("//button[text()= 'Start']")).Click();
            

            for (int l = 2; l <= 11; l++)
            {
                //Capturar informaçoes planilha
                var firstName = planilha.Cell($"A{l}").Value.ToString();
                var lastName = planilha.Cell($"B{l}").Value.ToString();
                var companyName = planilha.Cell($"C{l}").Value.ToString();
                var roleCompany = planilha.Cell($"D{l}").Value.ToString();
                var  Address = planilha.Cell($"E{l}").Value.ToString();
                var Email = planilha.Cell($"F{l}").Value.ToString();
                var Phone = planilha.Cell($"G{l}").Value.ToString();
               
                //input
                driver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelFirstName']")).SendKeys(firstName);
                driver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelLastName']")).SendKeys(lastName);
                driver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelCompanyName']")).SendKeys(companyName);
                driver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelRole']")).SendKeys(roleCompany);
                driver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelAddress']")).SendKeys(Address);
                driver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelEmail']")).SendKeys(Email);
                driver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelPhone']")).SendKeys(Phone);
                driver.FindElement(By.XPath("//input[@value= 'Submit']")).Click();

                Console.WriteLine($"{firstName} - {lastName} - {companyName} - {roleCompany} - {Address} - {Email} - {Phone}");

            }

            Console.Read();














        }


    }
}