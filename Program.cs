using System;
using FlaUI.Core;
using FlaUI.UIA3;
using FlaUI.UIA2;
using FlaUI.Core.Conditions;
using FlaUI.Core.AutomationElements;
using Serilog;
using System.Threading;
using System.Configuration;
using System.Runtime.InteropServices;


namespace DesktopDSPTTest // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        static Application appx;
        static Window DesktopWindow;
        static UIA3Automation automationUIA3 = new UIA3Automation();
        static ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());
        static AutomationElement window = automationUIA3.GetDesktop();
        static int step = 0;
        static string dtID = ConfigurationManager.AppSettings["dtID"];
        static string dtName = ConfigurationManager.AppSettings["dtName"];
        static string LoginId = ConfigurationManager.AppSettings["loginId"];
        static string LoginPassword = ConfigurationManager.AppSettings["password"];
        static string logfilename = ConfigurationManager.AppSettings["logfilename"];
        static string appfolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\" + ConfigurationManager.AppSettings["appfolder"];
        static string uploadfolder = appfolder + @"\" + ConfigurationManager.AppSettings["uploadfolder"];
        static string sharingfolder = appfolder + @"\" + ConfigurationManager.AppSettings["sharingfolder"];
        //static string enableconsolelog = appfolder + @"\" + ConfigurationManager.AppSettings["enableconsolelog"].ToUpper();
        //static string screenshotfolder = appfolder + @"\" + ConfigurationManager.AppSettings["screenshotfolder"];


        const UInt32 WM_CLOSE = 0x0010;

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        private static AutomationElement WaitForElement(Func<AutomationElement> findElementFunc)
        {
            AutomationElement element = null;
            for (int i = 0; i < 2000; i++)
            {
                element = findElementFunc();
                if (element != null)
                {
                    break;
                }

                Thread.Sleep(1);
            }
            return element;
        }
        static void closeWarningDialogBox(string Title)
        {
            var hWnd = FindWindow(null, Title);
            if (hWnd != IntPtr.Zero)
            {
                SendMessage(hWnd, WM_CLOSE, 0, 0);
                Log.Information("Closing dialog message box.");
            }
        }

        static void Main(string[] args)
        {
            try
            {
                if (!Directory.Exists(appfolder))
                {
                    Directory.CreateDirectory(appfolder);
                    Directory.CreateDirectory(uploadfolder);
                    Directory.CreateDirectory(sharingfolder);
                }
                var config = new LoggerConfiguration();
                config.WriteTo.Console();
                config.WriteTo.File(appfolder + logfilename);
                Log.Logger = config.CreateLogger();

                DeleteExcelFiles(uploadfolder);

                Log.Information("Accurate Desktop ver.4 Automation -  by FAIRBANC");


                if (!OpenAppAndDB())
                {
                    Log.Information("application automation failed !!");
                    return;
                }
                if (!LoginProcess())
                {
                    Log.Information("application automation failed !!");
                    return;
                }
                Log.Information("now wait for 1 minute before clicking report...");
                Thread.Sleep(35000);

                /* Try to navigare and open 'Sales' report */
                if (!OpenReport("sales"))
                {
                    Log.Information("Application Automation failed !!");
                    return;
                }
                /* Download opened report on screen */
                if (!DownloadReport("sales"))
                {
                    Log.Information("Application Automation failed !!");
                    return;
                }
                /* Closing Workspaces that contain all report tab */
                if (!ClosingWorkspace())
                {
                    Log.Information("Application Automation failed !!");
                    return;
                }
                /* Try to navigare and open 'Sales' report */
                if (!OpenReport("ar"))
                {
                    Log.Information("Application Automation failed !!");
                    return;
                }
                /* Download opened report on screen */
                if (!DownloadReport("ar"))
                {
                    Log.Information("Application Automation failed !!");
                    return;
                }
                /* Closing Workspaces that contain all report tab */
                if (!ClosingWorkspace())
                {
                    Log.Information("Application Automation failed !!");
                    return;
                }
            }
            catch (Exception ex)
            {
                Log.Information($"Error => {ex.ToString()}");
            }
            finally
            {
                if (automationUIA3 != null)
                { 
                    automationUIA3.Dispose();
                }
                 Log.CloseAndFlush();
            }
        }

        static bool ClosingWorkspace()
        {
            try
            {
                /* Travesing back to accurate desktop main windows */
                var eleMain = window.FindFirstDescendant(cf => cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                if (eleMain is null)
                {
                    Log.Information($"[Step #1 Quitting, end of ClosingWorkspace automation function !!");
                    return false;
                }
                eleMain.SetForeground();
                Thread.Sleep(1000);

                var ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                if (ele is null)
                {
                    Log.Information($"[Step #2 Quitting, end of OpenApp automation function !!");
                    return false;
                }
                Log.Information(ele.Properties.ClassName.ToString());
                ele.SetForeground();
                Log.Information(ele.Properties.ClassName.ToString());
                ele.Focus();
                Thread.Sleep(1000);

                System.Windows.Forms.SendKeys.SendWait("%o");
                Log.Information("Sending keys 'ALT+o'...");
                Thread.Sleep(1000);

                System.Windows.Forms.SendKeys.SendWait("A");
                Log.Information("Then sending key 'a'...");
                Thread.Sleep(1000);

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        static bool OpenAppAndDB()
        {
            try
            {
                appx = Application.Launch(@"C:\Program Files (x86)\CPSSoft\ACCURATE4 Enterprise\Accurate.exe");
                DesktopWindow = appx.GetMainWindow(automationUIA3);

                // Wait until Accurate window ready
                // FlaUI.Core.Input.Wait.UntilResponsive(DesktopWindow.FindFirstChild(), TimeSpan.FromMilliseconds(5000));
                Thread.Sleep(5000);
                closeWarningDialogBox("Welcome to Accurate");
                Log.Information("Action -> Closing 'Welcome to Accurate' window.");

                Thread.Sleep(1500);

                /* Closing Warning diaLog box */
                var ele3 = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Warning")));
                if (ele3 is null)
                {
                    Log.Information($"[Step #1] Quitting, end of OpenApp automation function !!");
                    return false;
                }
                ele3.FindFirstChild(cf => cf.ByName("OK")).AsButton().Click();

                var ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                if (ele is null)
                {
                    Log.Information($"[Step #2] Quitting, end of OpenApp automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.ClassName.ToString());
                ele.SetForeground();

                ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("File")));
                if (ele is null)
                {
                    Log.Information($"[Step #3] Quitting, end of OpenApp automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                ele.Click();

                // Context
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                if (ele is null)
                {
                    Log.Information("[Step #4] Quitting, end of OpenApp automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                //System.Windows.Forms.SendKeys.SendWait("%F");
                //Log.Information("Sending keys 'ALT+F'...");
                Thread.Sleep(1000);

                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(1);
                if (ele is null)
                {
                    Log.Information("[Step #5] Quitting, end of OpenApp automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named 'Context' with id -> " + ele.Properties.AutomationId.ToString());
                ele.Click();
                //System.Windows.Forms.SendKeys.SendWait("o");
                //Log.Information("Then sending key 'o'...");
                Thread.Sleep(1000);

                //Open Database
                ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Open Database")));
                if (ele is null)
                {
                    Log.Information($"[Step #6 Quitting, end of OpenApp automation function !!");
                    return false;
                }
                ele.Focus();

                var ele2 = WaitForElement(() => ele.FindFirstChild(cr => cr.ByClassName("TEdit")));
                if (ele2 is null)
                {
                    Log.Information($"[Step #7 Quitting, end of OpenApp automation function !!");
                    return false;
                }
                ele2.AsTextBox().Enter("C");
                if (ele2.AsTextBox().Text != @"C:\Program Files (x86)\CPSSoft\ACCURATE4 Enterprise\Sample\Sample.GDB")
                {
                    ///System.Windows.Forms.SendKeys.Send("{BACKSPACE}");
                    ele2.AsTextBox().Enter("\b \b");
                    ele2.AsTextBox().Enter(@"C:\Program Files (x86)\CPSSoft\ACCURATE4 Enterprise\Sample\Sample.GDB");
                }

                ele = ele.FindFirstChild(cf => cf.ByName("OK")).AsButton();
                if (ele2 is null)
                {
                    Log.Information($"[Step #8 Quitting, end of OpenApp automation function !!");
                    return false;
                }
                ele.Click();

                return true;
            }
            catch
            {
                Log.Information("Quitting, end of OpenApp automation function !!");
                return false;
            }
        }

        static bool LoginProcess()
        {
            try
            {
                var ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Login")));
                if (ele is null)
                {
                    Log.Information($"[Step #{step += 1}] Quitting, end of login automation function !!");
                    return false;
                }
                var ele2 = ele.FindFirstDescendant(cf => cf.ByClassName("TEdit")).AsTextBox();
                ele2.Enter(LoginId + "\t");
                System.Windows.Forms.SendKeys.SendWait(LoginPassword);
                ele.FindFirstDescendant(cf => cf.ByName("OK")).AsButton().Click();
                return true;
            }
            catch
            {
                Log.Information($"[Step #{step += 1}] Quitting, end of login automation function !!");
                return false;
            }
        }

        static bool OpenReport01(string rptType)
        {
            try
            {
                var ele1 = WaitForElement(() => window.FindFirstDescendant(cf => cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring)));
                if (ele1 is null)
                {
                    Log.Information($"[Step #1 Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                Log.Information("Class name is {0}", ele1.Properties.ClassName.ToString());
                ele1.SetForeground();
                Thread.Sleep(500);

                var ele = ele1.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar"));
                if (ele is null)
                {
                    Log.Information($"[Step #2] Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                Log.Information(ele.Properties.ClassName.ToString());
                //ele.Focus();

                System.Windows.Forms.SendKeys.SendWait("%R");
                Log.Information("Sending keys 'ALT+R'...");
                Thread.Sleep(1000);

                System.Windows.Forms.SendKeys.SendWait("i");
                Log.Information("Then sending key 'I'...");
                Thread.Sleep(1000);

                /* Travesing to 'Index to Reports' */
                //window = automationUIA3.GetDesktop();
                var eleParent = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Index to Reports")));
                if (eleParent is null)
                {
                    Log.Information($"[Step #3] Quitting, end of OpenReport0\\{rptType}1 automation function !!");
                    return false;
                }
                Log.Information("Class name is {0}", eleParent.Properties.ClassName.ToString());
                Thread.Sleep(1000);

                if (rptType == "sales")
                {
                    //Sales Reports
                    ele = eleParent.FindFirstDescendant(cf => cf.ByName("Sales Reports"));
                    if (ele is null)
                    {
                        Log.Information($"[Step #4] Quitting, end of OpenReport01\\{rptType} automation function !!");
                        return false;
                    }
                    ele.Click();
                    Thread.Sleep(500);

                    //Sales By Customer Detail
                    ele = eleParent.FindFirstDescendant(cf => cf.ByName("Sales By Customer Detail"));
                    if (ele is null)
                    {
                        Log.Information($"[Step #5] Quitting, end of OpenReport01\\{rptType} automation function !!"); ;
                        return false;
                    }
                    ele.DoubleClick();
                    Thread.Sleep(3000);
                }

                if (rptType == "ar")
                {
                    //Account Receivables & Customers
                    ele = eleParent.FindFirstDescendant(cf => cf.ByName("Account Receivables & Customers"));
                    if (ele is null)
                    {
                        Log.Information($"[Step #9] Quitting, end of OpenReport01\\{ rptType } automation function !!");
                        return false;
                    }
                    ele.Click();
                    Thread.Sleep(500);

                    //Invoices Paid Summary
                    ele = eleParent.FindFirstDescendant(cf => cf.ByName("Invoices Paid Summary"));
                    if (ele is null)
                    {
                        Log.Information($"[Step #10] Quitting, end of OpenReport01\\{rptType} automation function !!"); ;
                        return false;
                    }
                    ele.DoubleClick();
                    Thread.Sleep(3000);
                }

                //Report Format
                eleParent = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Report Format")).AsWindow());
                if (eleParent is null)
                {
                    Log.Information($"[Step #6] Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                Log.Information(eleParent.Properties.ClassName.ToString());
                eleParent.Focus();
                Thread.Sleep(2000);

                //  Filters && Parameters => Under 'Desktop' windows
                ele = eleParent.FindFirstDescendant(cf => cf.ByName("Filters && Parameters", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                //ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Filters && Parameters", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring)));
                if (ele is null)
                {
                    Log.Information($"[Step #7] Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                Log.Information(ele.Properties.ClassName.ToString());
                ele.Focus();
                Thread.Sleep(500);

                //TabDateFromTo
                ele = ele.FindFirstDescendant(cr => cf.ByName("TabDateFromTo"));
                if (ele is null)
                {
                    Log.Information($"[Step #8] Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                ele.Focus();

                /* Sending Report Date Parameters */
                AutomationElement[] ArrEle = ele.FindAllDescendants(cf => cf.ByClassName("TDateEdit"));
                if (ArrEle.Length > 0)
                {
                    //TDateEdit
                    Log.Information($"Number of DATE parameter on screen is : {ArrEle.Length}");
                    for (int index = ArrEle.Length - 1; index > -1; index--)
                    {
                        if (index != 0)
                        {
                            SendingDate(ArrEle[index], "01/01/2000");
                        }
                        else
                        {
                            SendingDate(ArrEle[index], "31/12/2023");
                        }
                    }
                }

                eleParent.FindFirstDescendant(cf.ByName("OK")).AsButton().Click();

                return true;
            }
            catch (Exception ex)
            {
                Log.Information($"[Exception] Quitting, end of OpenReport01\\{rptType} automation function !!");
                throw ex;
            }

        }

        static bool DownloadReport(string reportName)
        {
            try
            {
                /** Start downloading report process **/
                /* Travesing back to accurate desktop main windows */
                var ele1 = window.FindFirstDescendant(cf => cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                if (ele1 is null)
                {
                    Log.Information($"[Step #1 Quitting, end of DownloadReport automation function !!");
                    return false;
                }
                Log.Information("Class name is {0}", ele1.Properties.ClassName.ToString());
                Thread.Sleep(500);

                var ele = ele1.FindFirstDescendant(cf => cf.ByName("PriviewToolBar"));
                if (ele1 is null)
                {
                    Log.Information($"[Step #2 Quiting, end of DownloadReport automation function !!");
                    return false;
                }
                Log.Information("Class name is {0}", ele.Properties.ClassName.ToString());

                //Export settings
                ele.FindFirstChild(cf.ByName(("Export"))).AsButton().Click();
                Thread.Sleep(1000);


                /* The export button action resulting new window opened */
                ele1 = window.FindFirstDescendant(cf => cf.ByName("Export to Excel"));
                if (ele1 is null)
                {
                    Log.Information($"[Step #3 Quitting, end of DownloadReport automation function !!");
                    return false;
                }
                /* Put here the code for iteration of report parameter check box */
                /* End of codes */

                /* Clicking OK button  */
                ele1.FindFirstChild(cf => cf.ByName("OK")).AsButton().Click();

                if (!SavingFileDialog(reportName))
                { return false; }

                return true;
            }
            catch (Exception ex) { Log.Information(ex.ToString()); return false; }
        }

        private static bool SavingFileDialog(string reportName)
        {
            ///window = automationUIA3a.GetDesktop();
            //Save As
            //var ele1 = window.FindFirstDescendant(cf.ByClassName("#32770"));
            //if (ele1 is null)
            //{
            //    log.Information($"[Step #1] Quitting, end of OpenReport\\SavingFileDialog automation function !!");
            //    return false;
            //}
            //ele1.Focus();

            //Edit
            System.Windows.Forms.SendKeys.SendWait("%n");
            Log.Information("Sending keys 'ALT+n'...");
            Thread.Sleep(500);

            System.Windows.Forms.SendKeys.SendWait($@"{uploadfolder}\{reportName}.xls");
            Thread.Sleep(500);

            //Save
            var ele = window.FindFirstDescendant(cf => cf.ByName("Save"));
            if (ele is null)
            {
                Log.Information($"[Step #2] Quitting, end of OpenReport\\SavingFileDialog automation function !!");
                return false;
            }
            ele.AsButton().Click();

            /* Pause the app to wait file saving is finnished */
            Thread.Sleep(5000);

            return true;
        }

        private static bool SendingDate(AutomationElement ele, string date)
        {

            if (ele is null)
            {
                Log.Information($"[Step #1] Quitting, end of SendingDate automation function !!");
                return false;
            }

            ele.Click();

            // Send date parameter
            ele.AsTextBox().Enter("\b\b\b\b\b\b\b\b");
            ele.AsTextBox().Text = date;

            // TWinControl
            var childEle = ele.FindFirstDescendant(cf => cf.ByClassName("TWinControl"));
            if (childEle is null)
            {
                Log.Information($"[Step #2] Quitting, end of OpenReport01 automation function !!");
                return false;
            }

            childEle.Click();
            Thread.Sleep(500);
            childEle.Click();

            Log.Information($"Sending date parameter");

            return true;
        }

        private static void DeleteExcelFiles(string FolderName)
        {
            // Delete Excel files
            var supportFiles = Directory.EnumerateFiles(FolderName, "*.xl*");
            foreach (var excelFile in supportFiles)
            {
                File.Delete(excelFile);
                Log.Information($"Deleted Excel file: {excelFile}");
            }
        }

        static bool OpenReport(string reportType)
        {
            try
            {
                var mainElement = WaitForElement(() => window.FindFirstDescendant(cf.ByName("ACCURATE 4", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring)));
                if (mainElement is null)
                {
                    Log.Information($"[Step #1] Quitting, end of OpenReport automation function.");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + mainElement.Properties.Name.ToString());
                mainElement.SetForeground();
                Thread.Sleep(500);

                var ele = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                if (ele is null)
                {
                    Log.Information($"[Step #2] Quitting, end of OpenReport automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.ClassName.ToString());
                ele.SetForeground();
                Thread.Sleep(500);

                /* Click on Reports menu */
                ele = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByName("Reports")));
                if (ele is null)
                {
                    Log.Information($"[Step #3] Quitting, end of OpenReport automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                ele.AsMenu().Focus();
                ele.AsMenu().Click();
                Thread.Sleep(1000);

                // Context
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                if (ele is null)
                {
                    Log.Information("[Step #4] Quitting, end of OpenReport automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                //System.Windows.Forms.SendKeys.SendWait("%R");
                //Log.Information("Sending keys 'ALT+R'...");
                Thread.Sleep(1000);

                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(0);
                if (ele is null)
                {
                    Log.Information("[Step #5] Quitting, end of OpenReport automation function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named 'Context' with id -> " + ele.Properties.AutomationId.ToString());
                ele.Click();
                //System.Windows.Forms.SendKeys.SendWait("i");
                //Log.Information("Then sending key 'I'...");
                Thread.Sleep(2000);

                var indexToReportsElement = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Index to Reports")));
                if (indexToReportsElement == null)
                {
                    Log.Information($"[Step #6] Quitting, end of OpenReport OpenReport function.");
                    return false;
                }
                //indexToReportsElement.Click();
                Thread.Sleep(2000);

                var reportMain = (reportType == "sales") ? "Sales Reports" : "Account Receivables & Customers";
                var reportElement1 = indexToReportsElement.FindFirstDescendant(cf.ByName(reportMain));

                if (reportElement1 == null)
                {
                    Log.Information($"[Step #7] Quitting, end of OpenReport automation function.");
                    return false;
                }
                reportElement1.Click();
                Thread.Sleep(1000);

                var reportName = (reportType == "sales") ? "Sales By Customer Detail" : "Invoices Paid Summary";
                var reportElement2 = indexToReportsElement.FindFirstDescendant(cf.ByName(reportName));

                if (reportElement2 == null)
                {
                    Log.Information($"[Step #8] Quitting, end of OpenReport automation function.");
                    return false;
                }
                reportElement2.DoubleClick();
                Thread.Sleep(3000);

                // Report Format
                var reportFormatElement = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Report Format")).AsWindow());

                if (reportFormatElement == null)
                {
                    Log.Information($"[Step #9] Quitting, end of OpenReport automation function.");
                    return false;
                }

                reportFormatElement.Focus();
                Thread.Sleep(2000);

                // Filters && Parameters => Under 'Desktop' windows
                var filtersAndParametersElement = reportFormatElement.FindFirstDescendant(cf.ByName("Filters && Parameters", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));

                if (filtersAndParametersElement == null)
                {
                    Log.Information($"[Step #10] Quitting, end of OpenReport automation10unction.");
                    return false;
                }

                filtersAndParametersElement.Focus();
                Thread.Sleep(500);

                // TabDateFromTo
                var tabDateFromToElement = filtersAndParametersElement.FindFirstDescendant(cf.ByName("TabDateFromTo"));

                if (tabDateFromToElement == null)
                {
                    Log.Information($"[Step #11] Quitting, end of OpenReport aut[Step #11] ction.");
                    return false;
                }

                tabDateFromToElement.Focus();

                /* Sending Report Date Parameters */
                AutomationElement[] dateElements = tabDateFromToElement.FindAllDescendants(cf.ByClassName("TDateEdit"));

                if (dateElements.Length > 0)
                {
                    Log.Information($"Number of DATE parameters on screen is: {dateElements.Length}");

                    for (int index = dateElements.Length - 1; index > -1; index--)
                    {
                        if (index != 0)
                        {
                            SendingDate(dateElements[index], "01/01/2000");
                        }
                        else
                        {
                            SendingDate(dateElements[index], "31/12/2023");
                        }
                    }
                }

                reportFormatElement.FindFirstDescendant(cf.ByName("OK")).AsButton().Click();
                return true;
            }
            catch (Exception ex)
            {
                Log.Information($"Exception: {ex.ToString()}");
                return false;
            }
        }

    }

}
