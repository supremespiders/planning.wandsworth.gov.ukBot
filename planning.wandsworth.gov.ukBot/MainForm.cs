using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelHelperExe;
using MetroFramework.Controls;
using MetroFramework.Forms;
using Newtonsoft.Json;
using OpenQA.Selenium.Chrome;
using planning.wandsworth.gov.ukBot.Models;

namespace planning.wandsworth.gov.ukBot
{
    public partial class MainForm : MetroForm
    {
        public bool LogToUi = true;
        public bool LogToFile = true;

        private readonly string _path = Application.StartupPath;
        private int _maxConcurrency;
        private Dictionary<string, string> _config;
        public HttpCaller HttpCaller = new HttpCaller();
        Regex _regex = new Regex("[^a-zA-Z0-9]");
        TextInfo _textInfo = new CultureInfo("en-US", false).TextInfo;
        public MainForm()
        {
            InitializeComponent();
        }


        private async Task MainWork()
        {
            var url = await Task.Run(Search);
            url = url.Replace("&PS=10", "&PS=10000");
            NormalLog("Fetching links");
            var doc = await HttpCaller.GetDoc(url);
            doc.Save("doc.html");
            //var doc = new HtmlAgilityPack.HtmlDocument();
            //doc.LoadHtml(WebUtility.HtmlDecode(File.ReadAllText("doc.html")));
            var s = doc.DocumentNode.SelectSingleNode("//*[@id='lblPagePosition']");
            var nodes = doc.DocumentNode.SelectNodes("//a[@class='data_text']");
            var links = new List<string>();
            foreach (var node in nodes)
            {
                var href = node.GetAttributeValue("href", "").Replace("\t", "").Replace("\n", "").Replace("\r", "");
                links.Add("https://planning.wandsworth.gov.uk/Northgate/PlanningExplorer/Generic/" + href);
            }
            File.WriteAllLines("links", links);
            var items = await links.Work(_maxConcurrency, GetDetails);
            await items.SaveToExcel(outputI.Text);
            SuccessLog("work completed");
        }

        async Task<Item> GetDetails(string url)
        {
            var doc = await HttpCaller.GetDoc(url).ConfigureAwait(false);
            var nodes = doc.DocumentNode.SelectNodes("//h1[text()='Application Details']/following-sibling::ul/li/div/span");
            var item = new Item();
            var propertyInfo = typeof(Item).GetProperties().ToDictionary(x => x.Name);
            foreach (var node in nodes)
            {
                var name = _regex.Replace(_textInfo.ToTitleCase(node.InnerText), string.Empty);
                var value = node.NextSibling.InnerText.Replace("\r\t", "").Trim();
                propertyInfo[name].SetValue(item, value);
            }

            return item;
        }

        string Search()
        {
            ChromeDriver driver = null;
            try
            {
                NormalLog("Starting driver");
                var chromeDriverService = ChromeDriverService.CreateDefaultService();
                chromeDriverService.HideCommandPromptWindow = true;
                var options = new ChromeOptions();
                options.AddArgument("--window-position=-32000,-32000");
                driver = new ChromeDriver(chromeDriverService, options);
                NormalLog("Searching");
                driver.Navigate().GoToUrl("https://planning.wandsworth.gov.uk/Northgate/PlanningExplorer/GeneralSearch.aspx");
                driver.FindElementById("rbRange").Click();
                driver.FindElementById("dateStart").SendKeys(fromDateI.Value.ToString("d"));
                driver.FindElementById("dateEnd").SendKeys(toDateI.Value.ToString("d"));
                driver.FindElementById("csbtnSearch").Click();
                if (driver.FindElementsById("lblhtmlError").Count > 0)
                    throw new Exception($"Error on searching : {driver.FindElementById("lblhtmlError").Text}");
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
                var s = driver.FindElementById("lblPagePosition").Text;
                var url = driver.Url;
                Console.WriteLine(url);
                var sb = new StringBuilder();
                foreach (var cookiesAllCookie in driver.Manage().Cookies.AllCookies)
                    sb.Append($"{cookiesAllCookie.Name}={cookiesAllCookie.Value};");
                File.WriteAllText("ses", sb.ToString());
                HttpCaller = new HttpCaller();
                return url;
            }
            finally
            {
                driver?.Quit();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ServicePointManager.DefaultConnectionLimit = 65000;
            Application.ThreadException += Application_ThreadException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            Directory.CreateDirectory("data");
            outputI.Text = _path + @"\output.xlsx";
            LoadConfig();
            Utility.OnDisplay += OnDisplay;
            Utility.OnError += OnError;
        }

        private void OnError(object sender, string e)
        {
            ErrorLog(e);
        }

        private void OnDisplay(object sender, string e)
        {
            Display(e);
        }

        void InitControls(Control parent)
        {
            try
            {
                foreach (Control x in parent.Controls)
                {
                    try
                    {
                        if (x.Name.EndsWith("I"))
                        {
                            switch (x)
                            {
                                case MetroCheckBox _:
                                case CheckBox _:
                                    ((CheckBox)x).Checked = bool.Parse(_config[((CheckBox)x).Name]);
                                    break;
                                case RadioButton radioButton:
                                    radioButton.Checked = bool.Parse(_config[radioButton.Name]);
                                    break;
                                case TextBox _:
                                case RichTextBox _:
                                case MetroTextBox _:
                                    x.Text = _config[x.Name];
                                    break;
                                case MetroDateTime dateTime:
                                    dateTime.Value = DateTime.Parse(_config[dateTime.Name]);
                                    break;
                                case NumericUpDown numericUpDown:
                                    numericUpDown.Value = int.Parse(_config[numericUpDown.Name]);
                                    break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }

                    InitControls(x);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        public void SaveControls(Control parent)
        {
            try
            {
                foreach (Control x in parent.Controls)
                {
                    #region Add key value to disctionarry

                    if (x.Name.EndsWith("I"))
                    {
                        switch (x)
                        {
                            case MetroCheckBox _:
                            case RadioButton _:
                            case CheckBox _:
                                _config.Add(x.Name, ((CheckBox)x).Checked + "");
                                break;
                            case TextBox _:
                            case RichTextBox _:
                            case MetroTextBox _:
                                _config.Add(x.Name, x.Text);
                                break;
                            case MetroDateTime _:
                                _config.Add(x.Name, x.Text);
                                break;
                            case NumericUpDown _:
                                _config.Add(x.Name, ((NumericUpDown)x).Value + "");
                                break;
                            default:
                                Console.WriteLine(@"could not find a type for " + x.Name);
                                break;
                        }
                    }
                    #endregion
                    SaveControls(x);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        private void SaveConfig()
        {
            _config = new Dictionary<string, string>();
            SaveControls(this);
            try
            {
                File.WriteAllText("config.txt", JsonConvert.SerializeObject(_config, Formatting.Indented));
            }
            catch (Exception e)
            {
                ErrorLog(e.ToString());
            }
        }
        private void LoadConfig()
        {
            try
            {
                _config = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText("config.txt"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }
            InitControls(this);
        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.ToString(), @"Unhandled Thread Exception");
        }
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show((e.ExceptionObject as Exception)?.ToString(), @"Unhandled UI Exception");
        }
        #region UIFunctions
        public delegate void WriteToLogD(string s, Color c);
        public void WriteToLog(string s, Color c)
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new WriteToLogD(WriteToLog), s, c);
                    return;
                }
                if (LogToUi)
                {
                    if (DebugT.Lines.Length > 5000)
                    {
                        DebugT.Text = "";
                    }
                    DebugT.SelectionStart = DebugT.Text.Length;
                    DebugT.SelectionColor = c;
                    DebugT.AppendText(DateTime.Now.ToString(Utility.SimpleDateFormat) + " : " + s + Environment.NewLine);
                }
                Console.WriteLine(DateTime.Now.ToString(Utility.SimpleDateFormat) + @" : " + s);
                if (LogToFile)
                {
                    File.AppendAllText(_path + "/data/log.txt", DateTime.Now.ToString(Utility.SimpleDateFormat) + @" : " + s + Environment.NewLine);
                }
                if (c != Color.Red)
                    Display(s);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        public void NormalLog(string s)
        {
            WriteToLog(s, Color.Black);
        }
        public void ErrorLog(string s)
        {
            WriteToLog(s, Color.Red);
        }
        public void SuccessLog(string s)
        {
            WriteToLog(s, Color.Green);
        }
        public void CommandLog(string s)
        {
            WriteToLog(s, Color.Blue);
        }

        public delegate void SetProgressD(int x);
        public void SetProgress(int x)
        {
            if (InvokeRequired)
            {
                Invoke(new SetProgressD(SetProgress), x);
                return;
            }
            if ((x <= 100))
            {
                ProgressB.Value = x;
            }
        }
        public delegate void DisplayD(string s);
        public void Display(string s)
        {
            if (InvokeRequired)
            {
                Invoke(new DisplayD(Display), s);
                return;
            }
            displayT.Text = s;
        }

        #endregion
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveConfig();
        }

        private void openOutputB_Click_1(object sender, EventArgs e)
        {
            try
            {
                Process.Start(outputI.Text);
            }
            catch (Exception ex)
            {
                ErrorLog(ex.ToString());
            }
        }
        private void loadOutputB_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                Filter = @"xlsx file|*.xlsx",
                Title = @"Select the output location"
            };
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                outputI.Text = saveFileDialog1.FileName;
            }
        }

        private async void startB_Click_1(object sender, EventArgs e)
        {
            SaveConfig();
            LogToUi = logToUII.Checked;
            LogToFile = logToFileI.Checked;
            _maxConcurrency = (int)threadsI.Value;
            try
            {
                await MainWork();
            }
            catch (Exception exception)
            {
                ErrorLog(exception.ToString());
                Display(exception.Message);
            }
        }
    }
}
