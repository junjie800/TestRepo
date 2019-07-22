using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Text.RegularExpressions;

namespace UnitTestProjectTest
{
    [TestClass]
    public class Program
    {
        IWebDriver driver;
        static void Main(string[] args)
        { }
        [TestMethod]
        public void TheBingSearchTest()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://asknypadmindev.azurewebsites.net/botmain");
            driver.Manage().Window.Maximize();
            IWebElement imageclick = driver.FindElement(By.XPath("//img[@src='https://asknypadmin.azurewebsites.net/BotFolder/NYPChatBotRight.png']"));
            imageclick.Click();
            IWebElement frame = driver.FindElement(By.XPath(".//iframe[@id='nypBot']"));
            driver.SwitchTo().Frame(frame);
            driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/input")).Click();



            string questions;
            string columnheader;
            string newcolumnheaders;
            string answercells;
            string responsecells;
            int count = 0;
            int yescount = 0;
            excel.Application x1app = new excel.Application();
            x1app.SheetsInNewWorkbook = 1;
            x1app.Visible = true;
            excel.Workbook x1workbook = x1app.Workbooks.Open(@"C:\Users\L33539\Desktop\JUNJIE FYP PROJECT\Overall_QnA4");
            excel.Workbook NewWorkBook = x1app.Workbooks.Add();
            for (int x = 4; x <= 7; x++) //original x=4 x<=23
            {
                excel._Worksheet x1worksheet = x1workbook.Sheets[x];
                x1worksheet.Copy(Type.Missing, After: NewWorkBook.Sheets[x - 3]);
            }
            Thread.Sleep(2000);


            for (int w = 2; w <= 5; w++) //original w=2 w<=21
            {
                excel._Worksheet NewWorkSheet = NewWorkBook.Sheets[w];
                excel.Range NewWorkRange = NewWorkSheet.UsedRange;
                int colCount2 = NewWorkRange.Columns.Count;
                int rowCount2 = NewWorkRange.Rows.Count;
                NewWorkSheet.Cells[colCount2 + 1][1] = "Response";
                NewWorkSheet.Cells[colCount2 + 2][1] = "Timing of response retrieval";
                NewWorkSheet.Cells[colCount2 + 3][1] = "Does the answer and response match?";
                Console.WriteLine("Rows Count: " + (rowCount2 - 1));

                var numOfRes = 1;
                var newNumOfRes = 0;
                var testString = "";
                var testString2 = "";

                for (int i = 2; i <= rowCount2; i++)
                {

                    questions = NewWorkRange.Cells[1][i].value2;
                    driver.FindElement(By.XPath("//html/body/div[1]/div/div/div[3]/div/input")).SendKeys(questions); //"Send questions"
                    driver.FindElement(By.XPath("//html/body/div[1]/div/div/div[3]/button[1]")).Click(); //click button to send the question
                    Thread.Sleep(1000);

                    var textboxmsg = driver.FindElements(By.ClassName("format-markdown"));
                    var wrongqnmsg = driver.FindElements(By.XPath("//div[@class='wc-list']"));
                    count += 1;

                    newNumOfRes = textboxmsg.Count() + wrongqnmsg.Count();
                    if (newNumOfRes >= 1)
                    {
                        try
                        {
                            testString = textboxmsg.Last().GetAttribute("outerHTML");
                            Console.WriteLine("MESSAGES :" + testString);
                            testString2 = wrongqnmsg.Last().GetAttribute("outerHTML");
                        }
                        catch { }


                    }
                    while (newNumOfRes == numOfRes)
                    {
                        // Pooling
                        Thread.Sleep(1000);
                        textboxmsg = driver.FindElements(By.ClassName("format-markdown"));
                        wrongqnmsg = driver.FindElements(By.XPath("//div[@class='wc-list']"));
                        newNumOfRes = textboxmsg.Count() + wrongqnmsg.Count();
                        try
                        {
                            testString = textboxmsg.Last().GetAttribute("outerHTML");
                            testString2 = wrongqnmsg.Last().GetAttribute("outerHTML");
                        }
                        catch
                        { }
                    }

                    numOfRes = newNumOfRes;
                    // foreach (var textmsg in textboxmsg)
                    // {
                    for (int c = 1; c <= colCount2 + 2; c++)
                    {

                        columnheader = NewWorkRange.Cells[1, c].value2;
                        newcolumnheaders = NewWorkSheet.Cells[1, c].value2;
                        answercells = NewWorkSheet.Cells[i, 2].Text;
                        responsecells = NewWorkSheet.Cells[i, 3].Text;

                        //retrieve response with all tags then remove all the tags below
                        //try
                        //{
                        var outerhtml = testString; //figure this out later (system no element exception)
                        //}
                        //catch
                        //{

                        //}

                        outerhtml = outerhtml.Replace("<br />", Environment.NewLine);
                        outerhtml = Regex.Replace(outerhtml, @"<(?!a|/a|ol|ul[\x20/>])[^<>]+>", string.Empty);
                        outerhtml = outerhtml.TrimEnd('\r', '\n');  //remove carriage return
                                                                    //all to replace some to match
                        outerhtml = outerhtml.Replace("“", "\"");
                        outerhtml = outerhtml.Replace("”", "\"");
                        outerhtml = outerhtml.Replace("<ul>", "-");
                        outerhtml = outerhtml.Replace("‘", "'");
                        outerhtml = outerhtml.Replace("’", "'");

                        var outerhtml2 = wrongqnmsg.Last().GetAttribute("outerHTML");
                        outerhtml2 = outerhtml2.Replace("<br />", Environment.NewLine);
                        outerhtml2 = Regex.Replace(outerhtml2, @"<(?!a|/a|ol|ul[\x20/>])[^<>]+>", string.Empty);
                        outerhtml2 = outerhtml2.TrimEnd('\r', '\n');  //remove carriage return
                                                                      //all to replace some to match
                        outerhtml2 = outerhtml2.Replace("“", "\"");
                        outerhtml2 = outerhtml2.Replace("”", "\"");
                        outerhtml2 = outerhtml2.Replace("<ul>", "-");
                        outerhtml2 = outerhtml2.Replace("‘", "'");
                        outerhtml2 = outerhtml2.Replace("’", "'");


                        //to replace ol with numerics
                        int result = 0;
                        StringBuilder sb = new StringBuilder(outerhtml);
                        result = outerhtml.IndexOf("<ol");
                        while (result > -1)
                        {
                            if (result == outerhtml.IndexOf("<ol>"))
                            {
                                sb.Remove(result, 4);
                                sb.Insert(result, "1)");
                            }
                            else
                            {
                                char number = outerhtml[result + 11];
                                sb.Remove(result, 14);
                                sb.Insert(result, number + ")");

                            }
                            outerhtml = sb.ToString();
                            result = outerhtml.IndexOf("<ol");
                        }

                        //below is to remove linebreaks and whitespace for both answer and response cells to do matching
                        var compareresponsecells = Regex.Replace(outerhtml, @"\r\n?|\n", ""); //to remove line breaks for comparison
                        compareresponsecells = Regex.Replace(compareresponsecells, @"\s+", ""); //to remove whitespace for comparison
                        var compareanswercells = Regex.Replace(answercells, @"\r\n?|\n", "");
                        compareanswercells = Regex.Replace(compareanswercells, @"\s+", "");

                        //Console.WriteLine(newcolumnheaders);
                        if (columnheader == "Question")
                        {

                        }
                        else if (columnheader == "Answer")
                        {

                        }
                        else if (columnheader == "Answers")
                        {

                        }
                        else if (newcolumnheaders == "Response")
                        {
                            //Console.WriteLine("YES " + count);
                            try
                            {
                                NewWorkSheet.Cells[i, c] = outerhtml;
                                Console.WriteLine("SICK FEELING :" + outerhtml);
                                var wcmessagecontented = driver.FindElements(By.XPath("//div[@class='wc-message-content']"));
                                var lastwcmsgcontent = wcmessagecontented.Last();
                                var child = lastwcmsgcontent.FindElement(By.XPath("./div/div")); // ./ means go down from this element
                                //Console.WriteLine("LAST CHILD :" + child.GetAttribute("outerHTML"));
                                if (child.GetAttribute("class").Contains("wc-list"))
                                {
                                    NewWorkSheet.Cells[i, c] = outerhtml2;
                                    //Console.WriteLine("Panini");
                                }
                                //NewWorkSheet.Cells[i, c] = outerhtml2;
                            }
                            catch
                            {
                            }
                            //outerhtml.Contains((char)13);
                            //Console.WriteLine(outerhtml.Contains((char)13));
                            //Console.WriteLine("WELP:" + outerhtml);

                        }
                        else if (newcolumnheaders == "Timing of response retrieval")
                        {
                            try
                            {
                                NewWorkSheet.Cells[i, c] = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");
                            }
                            catch
                            {
                            }
                        }
                        else if (newcolumnheaders == "Does the answer and response match?")
                        {
                            if (compareanswercells.Equals(compareresponsecells))
                            {
                                try
                                {
                                    NewWorkSheet.Cells[i, c] = "Yes";
                                }
                                catch
                                {

                                }
                            }
                            else
                            {
                                try
                                {
                                    NewWorkSheet.Cells[i, c] = "No";
                                }
                                catch
                                {
                                }
                            }
                        }
                        else if (columnheader == null)
                        {

                        }
                        else
                        {
                            try
                            {
                                NewWorkSheet.Columns[c].Delete();
                                c--;
                            }
                            catch
                            {

                            }
                        }
                    }
                    // }
                    if (NewWorkSheet.Cells[i, 5].Text == "Yes")
                    {
                        yescount += 1;
                        //Console.WriteLine("For yes:" + NewWorkSheet.Cells[1][i].Text);
                        //Console.WriteLine("Yes: " + yescount);
                    }
                    else
                    {

                    }
                    //Console.WriteLine("TOTALMATCHCOUNT:" + count);

                    //x1workbook.SaveAs(path, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();


                }
                excel._Worksheet NewWorkSheet1 = NewWorkBook.Sheets[1];
                excel.Range NewWorkRange1 = NewWorkSheet1.UsedRange;
                //int colCount3 = NewWorkRange.Columns.Count;
                //int rowCount3 = NewWorkRange.Rows.Count;
                NewWorkSheet1.Cells[1][1] = "Total Count of Matches: " + count;
                NewWorkSheet1.Cells[1][2] = "Total Count of Matches Matched: " + yescount;
            }

        }



    }
}
