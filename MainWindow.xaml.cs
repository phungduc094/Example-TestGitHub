using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace SeleniumBHXH
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }


        #region Case 1: Test BHXH Tu Nguyen

        private void Button_Click1(object sender, RoutedEventArgs e)
        {
            string path = @"D:\Documents\SQA\BTL\Selenium\BHXH_Tu_Nguyen.xlsx";
            int sheet = 1;

            _Application excel = new _Excel.Application();
            Workbook wb = excel.Workbooks.Open(path);
            Worksheet ws = wb.Worksheets[sheet];

            int startRow = 3, endRow = 11;
            List<string> inputs1 = getInputs(ws, startRow, endRow, 3);
            List<string> inputs2 = getInputs(ws, startRow, endRow, 4);
            List<string> inputs3 = getInputs(ws, startRow, endRow, 5);
            List<string> inputs4 = getInputs(ws, startRow, endRow, 6);
            List<string> inputs5 = getInputs(ws, startRow, endRow, 7);
            List<string> results = getInputs(ws, startRow, endRow, 8);

            #region Selenium

            ChromeDriver chromeDriver = new ChromeDriver();

            // open a web
            chromeDriver.Url = "http://localhost:3000/bhxh/tu-nguyen";
            chromeDriver.Navigate();

            chromeDriver.Manage().Window.Maximize();

            int resultCol = 9, stateCol = 10;
            string pathImage = @"D:\Documents\SQA\BTL\Selenium\case";
            for (int i = 0; i < inputs1.Count; i++)
            {
                chromeDriver.Navigate().Refresh();
                int currentRow = i + startRow;

                var input1 = chromeDriver.FindElements(By.XPath("//input[@class='form-control']"))[0]; // CMTND
                input1.SendKeys(inputs1[i]);

                var input2 = chromeDriver.FindElements(By.XPath("//input[@class='form-control']"))[1]; // Ho ten
                input2.SendKeys(inputs2[i]);

                var input3 = chromeDriver.FindElements(By.XPath("//input[@class='form-control']"))[2]; // Dia chi
                input3.SendKeys(inputs3[i]);

                var input4 = chromeDriver.FindElements(By.XPath("//input[@class='form-control']"))[3]; // Loai bao hien
                input4.SendKeys(inputs4[i]);

                var input5 = chromeDriver.FindElements(By.XPath("//input[@class='form-control form-control']"))[0];
                input5.Clear();
                input5.SendKeys(inputs5[i]);

                var submit = chromeDriver.FindElements(By.XPath("//button[@class='btn btn-primary']"))[0];
                submit.Click();

                Screenshot saveScreenShot = ((ITakesScreenshot)chromeDriver).GetScreenshot();
                saveScreenShot.SaveAsFile(pathImage + (i + 1).ToString() + ".png", ScreenshotImageFormat.Png);

                List<IWebElement> messages = new List<IWebElement>();
                messages.AddRange(chromeDriver.FindElements(By.XPath("//div[@class='validation-error mb-3']")));
                if (messages.Count == 0)
                {
                    WriteToCell(ws, currentRow, resultCol, "Success");
                    WriteToCell(ws, currentRow, stateCol, "PASSED");
                }
                else if (messages.Count == 1)
                {
                    string messageText = messages[0].Text;

                    WriteToCell(ws, currentRow, resultCol, messageText);
                    if (messageText == results[i]) WriteToCell(ws, currentRow, stateCol, "PASSED");
                    else WriteToCell(ws, currentRow, stateCol, "FALSE");
                }
            }

            #endregion

            wb.Close();
            excel.Quit();
        }

        private List<string> getInputs(Worksheet ws, int startRow, int endRow, int col)
        {
            List<string> list = new List<string>();
            
            for(int i = startRow; i <= endRow; i++)
            {
                string tmp = "";
                if (ws.Cells[i, col].Value != null)
                {
                    tmp = ws.Cells[i, col].Value.ToString();
                }

                list.Add(tmp);
            }

            return list;
        }

        private void WriteToCell(Worksheet ws, int row, int col, string value)
        {
            ws.Cells[row, col].Value = value;
        }

        #endregion

        #region Case 2: BHXH Bat Buoc

        private void Button_Click2(object sender, RoutedEventArgs e)
        {

        }

        #endregion
    }
}