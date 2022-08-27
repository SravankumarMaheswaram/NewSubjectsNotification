using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using OpenQA.Selenium.Support.UI;
using System.Net;
using System.Net.Mail;
using System.Configuration;

namespace ConsoleApp2
{

    class Program
    {
        static IWebDriver driver;

        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            driver = new ChromeDriver(ConfigurationManager.AppSettings.Get("driverpath"));

            driver.Url = "https://elion.saintleo.edu/";
            System.Threading.Thread.Sleep(5000);
            bool boolIsData = false;

            IList<IWebElement> elements = driver.FindElements(By.Id("acctLogin"));
            foreach (IWebElement e in elements)
            {
                e.Click();
                //USER_NAME
                driver.FindElement(By.Id("USER_NAME")).SendKeys("sravankumar.mahesw");

                driver.FindElement(By.Id("CURR_PWD")).SendKeys("Welcome2022$");
                driver.FindElement(By.Name("SUBMIT2")).Click();
                driver.FindElement(By.ClassName("WBST_Bars")).Click();
                driver.FindElement(By.LinkText("Search/Register for Sections")).Click();

                IList<IWebElement> elementsitem = driver.FindElements(By.Id("VAR1"));
                foreach (IWebElement item in elementsitem)
                {
                    Console.WriteLine(item.Text);
                }
                IWebElement selectElement = driver.FindElement(By.Id("VAR1"));
                var selectObject = new SelectElement(selectElement);
                selectObject.SelectByValue("2022FA1");

                IWebElement selectEleSub = driver.FindElement(By.Id("LIST_VAR1_1"));
                var selectObjSub = new SelectElement(selectEleSub);
                selectObjSub.SelectByValue("COM");

                IWebElement selectEleCL = driver.FindElement(By.Id("LIST_VAR2_1"));
                var selectObjCL = new SelectElement(selectEleCL);
                selectObjCL.SelectByValue("500");

                driver.FindElement(By.Name("SUBMIT2")).Click();
                int i = 0;
                List<string> subjects = new List<string>();
                string localtext = string.Empty;
                try
                {
                    do
                    {
                        i++;
                        localtext = string.Concat(driver.FindElement(By.Id("SEC_SHORT_TITLE_" + i.ToString())).Text + " ~",
                            driver.FindElement(By.Id("SEC_FACULTY_INFO_" + i.ToString())).Text + " ~",
                            driver.FindElement(By.Id("SEC_MEETING_INFO_" + i.ToString())).Text + " ~",
                            driver.FindElement(By.Id("LIST_VAR3_" + i.ToString())).Text);

                        string subjectcodes = ConfigurationManager.AppSettings.Get("SubjectCodes");
                        string[] subcode = subjectcodes.Split(",");
                        foreach (var sub in subcode)
                        {
                            if (localtext.Contains(sub))
                            {
                                subjects.Add(localtext);
                                boolIsData = true;
                            }
                        }

                        //if (localtext.Contains("542") || localtext.Contains("562") || localtext.Contains("560") || localtext.Contains("571") || localtext.Contains("575"))
                        //    subjects.Add(localtext);


                    } while (i <= 100);
                }
                catch 
                {
                    Console.WriteLine("data Added");
                }

                foreach (var item in subjects)
                {
                    Console.WriteLine(item);
                }

                string mailbody = string.Empty;
                try
                {
                    if (subjects.Count != 0)
                    {
                        string messageBody = "<font>The following are the cources: </font><br><br>";
                        string htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
                        string htmlTableEnd = "</table>";
                        string htmlHeaderRowStart = "<tr style=\"background-color:#6FA1D2; color:#ffffff;\">";
                        string htmlHeaderRowEnd = "</tr>";
                        string htmlTrStart = "<tr style=\"color:#555555;\">";
                        string htmlTrEnd = "</tr>";
                        string htmlTdStart = "<td style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;\">";
                        string htmlTdEnd = "</td>";
                        messageBody += htmlTableStart;
                        messageBody += htmlHeaderRowStart;
                        messageBody += htmlTdStart + "Subject Name" + htmlTdEnd;
                        messageBody += htmlTdStart + "Professor" + htmlTdEnd;
                        messageBody += htmlTdStart + "Class Info" + htmlTdEnd;
                        messageBody += htmlTdStart + "Seats Available" + htmlTdEnd;
                        messageBody += htmlHeaderRowEnd;
                        //Loop all the rows from grid vew and added to html td  
                        for (int j = 0; j <= subjects.Count - 1; j++)
                        {
                            messageBody = messageBody + htmlTrStart;
                            messageBody = messageBody + htmlTdStart + subjects[j].Split("~")[0] + htmlTdEnd; //adding student name  
                            messageBody = messageBody + htmlTdStart + subjects[j].Split("~")[1] + htmlTdEnd; //adding DOB  
                            messageBody = messageBody + htmlTdStart + subjects[j].Split("~")[2] + htmlTdEnd; //adding DOB  
                            messageBody = messageBody + htmlTdStart + subjects[j].Split("~")[3] + htmlTdEnd; //adding Email  
                            messageBody = messageBody + htmlTrEnd;
                        }
                        messageBody = messageBody + htmlTableEnd;
                        // return HTML Table as string from this function  
                        mailbody += messageBody + "<br>";


                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                try
                {
                    //SrJ
                    //driver.Url = "https://elion.saintleo.edu/";

                    driver.FindElement(By.LinkText("Students Menu")).Click();
                    driver.FindElement(By.LinkText("Search/Register for Sections")).Click();

                    IList<IWebElement> elementsitem1 = driver.FindElements(By.Id("VAR1"));
                    IWebElement selectElement1 = driver.FindElement(By.Id("VAR1"));
                    var selectObject1 = new SelectElement(selectElement1);
                    selectObject1.SelectByValue("2022FA2");

                    IWebElement selectEleSub1 = driver.FindElement(By.Id("LIST_VAR1_1"));
                    var selectObjSub1 = new SelectElement(selectEleSub1);
                    selectObjSub1.SelectByValue("COM");

                    IWebElement selectEleCL1 = driver.FindElement(By.Id("LIST_VAR2_1"));
                    var selectObjCL1 = new SelectElement(selectEleCL1);
                    selectObjCL1.SelectByValue("500");

                    driver.FindElement(By.Name("SUBMIT2")).Click();
                    i = 0;
                    List<string> subjects1 = new List<string>();
                    localtext = string.Empty;
                    try
                    {
                        do
                        {
                            i++;
                            localtext = string.Concat(driver.FindElement(By.Id("SEC_SHORT_TITLE_" + i.ToString())).Text + " ~",
                                driver.FindElement(By.Id("SEC_FACULTY_INFO_" + i.ToString())).Text + " ~",
                                driver.FindElement(By.Id("SEC_MEETING_INFO_" + i.ToString())).Text + " ~",
                                driver.FindElement(By.Id("LIST_VAR3_" + i.ToString())).Text);

                            string subjectcodes = ConfigurationManager.AppSettings.Get("SubjectCodes");
                            string[] subcode = subjectcodes.Split(",");
                            foreach (var sub in subcode)
                            {
                                if (localtext.Contains(sub))
                                {
                                    subjects1.Add(localtext);
                                    boolIsData = true;
                                }
                            }

                            //if (localtext.Contains("542") || localtext.Contains("562") || localtext.Contains("560") || localtext.Contains("571") || localtext.Contains("575"))
                            //    subjects1.Add(localtext);


                        } while (i <= 100);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(" Ex data Added 1");
                        //ignore
                    }

                    foreach (var item in subjects)
                    {
                        Console.WriteLine(item);
                    }

                    string messageBody = "<br><p> Fall 2 subjects </p>";// "<font>The following are the cources: </font><br><br>";

                    string htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
                    string htmlTableEnd = "</table>";
                    string htmlHeaderRowStart = "<tr style=\"background-color:#6FA1D2; color:#ffffff;\">";
                    string htmlHeaderRowEnd = "</tr>";
                    string htmlTrStart = "<tr style=\"color:#555555;\">";
                    string htmlTrEnd = "</tr>";
                    string htmlTdStart = "<td style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;\">";
                    string htmlTdEnd = "</td>";
                    messageBody += htmlTableStart;
                    messageBody += htmlHeaderRowStart;
                    messageBody += htmlTdStart + "Subject Name" + htmlTdEnd;
                    messageBody += htmlTdStart + "Professor" + htmlTdEnd;
                    messageBody += htmlTdStart + "Class Info" + htmlTdEnd;
                    messageBody += htmlTdStart + "Seats Available" + htmlTdEnd;
                    messageBody += htmlHeaderRowEnd;
                    //Loop all the rows from grid vew and added to html td  
                    for (int j = 0; j <= subjects1.Count - 1; j++)
                    {
                        messageBody = messageBody + htmlTrStart;
                        messageBody = messageBody + htmlTdStart + subjects1[j].Split("~")[0] + htmlTdEnd; //adding student name  
                        messageBody = messageBody + htmlTdStart + subjects1[j].Split("~")[1] + htmlTdEnd; //adding DOB  
                        messageBody = messageBody + htmlTdStart + subjects1[j].Split("~")[2] + htmlTdEnd; //adding DOB  
                        messageBody = messageBody + htmlTdStart + subjects1[j].Split("~")[3] + htmlTdEnd; //adding Email  
                        messageBody = messageBody + htmlTrEnd;
                    }
                    messageBody = messageBody + htmlTableEnd;
                    // return HTML Table as string from this function  
                    mailbody += messageBody;

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ex data added");
                }

                try
                {
                    //SrJ1
                    //driver.Url = "https://elion.saintleo.edu/";

                    driver.FindElement(By.LinkText("Students Menu")).Click();
                    driver.FindElement(By.LinkText("Search/Register for Sections")).Click();

                    IList<IWebElement> elementsitem1 = driver.FindElements(By.Id("VAR1"));

                    IWebElement selectElement1 = driver.FindElement(By.Id("VAR1"));
                    var selectObject1 = new SelectElement(selectElement1);
                    selectObject1.SelectByValue("2023SP1");

                    IWebElement selectEleSub1 = driver.FindElement(By.Id("LIST_VAR1_1"));
                    var selectObjSub1 = new SelectElement(selectEleSub1);
                    selectObjSub1.SelectByValue("COM");

                    IWebElement selectEleCL1 = driver.FindElement(By.Id("LIST_VAR2_1"));
                    var selectObjCL1 = new SelectElement(selectEleCL1);
                    selectObjCL1.SelectByValue("500");

                    driver.FindElement(By.Name("SUBMIT2")).Click();
                    i = 0;
                    List<string> subjects2 = new List<string>();
                    localtext = string.Empty;
                    try
                    {
                        do
                        {
                            i++;
                            localtext = string.Concat(driver.FindElement(By.Id("SEC_SHORT_TITLE_" + i.ToString())).Text + " ~",
                                driver.FindElement(By.Id("SEC_FACULTY_INFO_" + i.ToString())).Text + " ~",
                                driver.FindElement(By.Id("SEC_MEETING_INFO_" + i.ToString())).Text + " ~",
                                driver.FindElement(By.Id("LIST_VAR3_" + i.ToString())).Text);

                            string subjectcodes = ConfigurationManager.AppSettings.Get("SubjectCodes");
                            string[] subcode = subjectcodes.Split(",");
                            foreach (var sub in subcode)
                            {
                                if (localtext.Contains(sub))
                                {
                                    subjects2.Add(localtext);
                                    boolIsData = true;
                                }
                            }

                            //if (localtext.Contains("542") || localtext.Contains("562") || localtext.Contains("560") || localtext.Contains("571") || localtext.Contains("575"))
                            //    subjects2.Add(localtext);

                        } while (i <= 100);
                    }
                    catch 
                    {
                        Console.WriteLine(" Ex data Added 1");
                    }

                    foreach (var item in subjects)
                    {
                        Console.WriteLine(item);
                    }

                    string messageBody = "<br><p> spring 1 subjects </p>";// "<font>The following are the cources: </font><br><br>";

                    string htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
                    string htmlTableEnd = "</table>";
                    string htmlHeaderRowStart = "<tr style=\"background-color:#6FA1D2; color:#ffffff;\">";
                    string htmlHeaderRowEnd = "</tr>";
                    string htmlTrStart = "<tr style=\"color:#555555;\">";
                    string htmlTrEnd = "</tr>";
                    string htmlTdStart = "<td style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;\">";
                    string htmlTdEnd = "</td>";
                    messageBody += htmlTableStart;
                    messageBody += htmlHeaderRowStart;
                    messageBody += htmlTdStart + "Subject Name" + htmlTdEnd;
                    messageBody += htmlTdStart + "Professor" + htmlTdEnd;
                    messageBody += htmlTdStart + "Class Info" + htmlTdEnd;
                    messageBody += htmlTdStart + "Seats Available" + htmlTdEnd;
                    messageBody += htmlHeaderRowEnd;
                    //Loop all the rows from grid vew and added to html td  
                    for (int j = 0; j <= subjects2.Count - 1; j++)
                    {
                        messageBody = messageBody + htmlTrStart;
                        messageBody = messageBody + htmlTdStart + subjects2[j].Split("~")[0] + htmlTdEnd; //adding student name  
                        messageBody = messageBody + htmlTdStart + subjects2[j].Split("~")[1] + htmlTdEnd; //adding DOB
                        messageBody = messageBody + htmlTdStart + subjects2[j].Split("~")[2] + htmlTdEnd; //adding DOB  
                        messageBody = messageBody + htmlTdStart + subjects2[j].Split("~")[3] + htmlTdEnd; //adding Email  
                        messageBody = messageBody + htmlTrEnd;
                    }
                    messageBody = messageBody + htmlTableEnd;
                    // return HTML Table as string from this function  
                    mailbody += messageBody;

                }
                catch 
                {
                    Console.WriteLine("Ex data added");
                }

                try
                {
                    //SrJ1
                    //driver.Url = "https://elion.saintleo.edu/";

                    driver.FindElement(By.LinkText("Students Menu")).Click();
                    driver.FindElement(By.LinkText("Search/Register for Sections")).Click();

                    IList<IWebElement> elementsitem1 = driver.FindElements(By.Id("VAR1"));

                    IWebElement selectElement1 = driver.FindElement(By.Id("VAR1"));
                    var selectObject1 = new SelectElement(selectElement1);
                    selectObject1.SelectByValue("2023SP2");

                    IWebElement selectEleSub1 = driver.FindElement(By.Id("LIST_VAR1_1"));
                    var selectObjSub1 = new SelectElement(selectEleSub1);
                    selectObjSub1.SelectByValue("COM");

                    IWebElement selectEleCL1 = driver.FindElement(By.Id("LIST_VAR2_1"));
                    var selectObjCL1 = new SelectElement(selectEleCL1);
                    selectObjCL1.SelectByValue("500");

                    driver.FindElement(By.Name("SUBMIT2")).Click();
                    i = 0;
                    List<string> subjects3 = new List<string>();
                    localtext = string.Empty;
                    try
                    {
                        do
                        {
                            i++;
                            localtext = string.Concat(driver.FindElement(By.Id("SEC_SHORT_TITLE_" + i.ToString())).Text + " ~",
                                driver.FindElement(By.Id("SEC_FACULTY_INFO_" + i.ToString())).Text + " ~",
                                driver.FindElement(By.Id("SEC_MEETING_INFO_" + i.ToString())).Text + " ~",
                                driver.FindElement(By.Id("LIST_VAR3_" + i.ToString())).Text);

                            string subjectcodes = ConfigurationManager.AppSettings.Get("SubjectCodes");
                            string[] subcode = subjectcodes.Split(",");
                            foreach (var sub in subcode)
                            {
                                if (localtext.Contains(sub))
                                {
                                    subjects3.Add(localtext);
                                    boolIsData = true;
                                }
                            }

                            //if (localtext.Contains("542") || localtext.Contains("562") || localtext.Contains("560") || localtext.Contains("571") || localtext.Contains("575"))
                            //    subjects3.Add(localtext);

                        } while (i <= 100);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(" Ex data Added 1");
                    }

                    foreach (var item in subjects)
                    {
                        Console.WriteLine(item);
                    }

                    string messageBody = "<br><p> spring 2 subjects </p>";// "<font>The following are the cources: </font><br><br>";
                    string htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
                    string htmlTableEnd = "</table>";
                    string htmlHeaderRowStart = "<tr style=\"background-color:#6FA1D2; color:#ffffff;\">";
                    string htmlHeaderRowEnd = "</tr>";
                    string htmlTrStart = "<tr style=\"color:#555555;\">";
                    string htmlTrEnd = "</tr>";
                    string htmlTdStart = "<td style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;\">";
                    string htmlTdEnd = "</td>";
                    messageBody += htmlTableStart;
                    messageBody += htmlHeaderRowStart;
                    messageBody += htmlTdStart + "Subject Name" + htmlTdEnd;
                    messageBody += htmlTdStart + "Professor" + htmlTdEnd;
                    messageBody += htmlTdStart + "Class Info" + htmlTdEnd;
                    messageBody += htmlTdStart + "Seats Available" + htmlTdEnd;
                    messageBody += htmlHeaderRowEnd;
                    //Loop all the rows from grid vew and added to html td  
                    for (int j = 0; j <= subjects.Count - 1; j++)
                    {
                        messageBody = messageBody + htmlTrStart;
                        messageBody = messageBody + htmlTdStart + subjects3[j].Split("~")[0] + htmlTdEnd; //adding student name  
                        messageBody = messageBody + htmlTdStart + subjects3[j].Split("~")[1] + htmlTdEnd; //adding DOB  
                        messageBody = messageBody + htmlTdStart + subjects3[j].Split("~")[2] + htmlTdEnd; //adding DOB  
                        messageBody = messageBody + htmlTdStart + subjects3[j].Split("~")[3] + htmlTdEnd; //adding Email  
                        messageBody = messageBody + htmlTrEnd;
                    }
                    messageBody = messageBody + htmlTableEnd;
                    // return HTML Table as string from this function  
                    mailbody += messageBody;

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ex data added");
                }

                if (boolIsData)
                {
                    //send mail
                    MailMessage message = new MailMessage();
                    SmtpClient smtp = new SmtpClient();
                    message.From = new MailAddress(ConfigurationManager.AppSettings.Get("FromMail"));// "msravankumar36@outlook.com"
                    message.To.Add(new MailAddress(ConfigurationManager.AppSettings.Get("ToMail"))); ;// "msravankumar36@outlook.com"
                    message.To.Add(new MailAddress(ConfigurationManager.AppSettings.Get("ToMail1")));//"sravankumar.mahesw@email.saintleo.edu"
                    //message.To.Add(new MailAddress("giriprasad.surinen@email.saintleo.edu"));
                    //message.To.Add(new MailAddress("bindu.gadipally@email.saintleo.edu"));

                    message.Subject = ConfigurationManager.AppSettings.Get("subject");//"Subjects"
                    message.IsBodyHtml = true; //to make message body as html  
                    message.Body = mailbody;
                    smtp.Port = Convert.ToInt32(ConfigurationManager.AppSettings.Get("port"));// 25;//587;
                    smtp.Host = ConfigurationManager.AppSettings.Get("host");// "smtp-mail.outlook.com";//"smtp.gmail.com"; //for gmail host  

                    smtp.UseDefaultCredentials = false;
                    smtp.EnableSsl = true;
                    smtp.Credentials = new NetworkCredential(ConfigurationManager.AppSettings.Get("username"), ConfigurationManager.AppSettings.Get("password"));

                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    Console.WriteLine("Email prepared");
                    smtp.Send(message);
                    Console.WriteLine("Email sent");
                }
                else
                {
                    Console.WriteLine("no new subjects");
                }
            }

            //System.Threading.Thread.Sleep(5000);
            driver.Close();
        }
    }
}
