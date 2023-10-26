using System;
using FlaUI.Core;
using FlaUI.UIA3;
using FlaUI.UIA2;
using FlaUI.Core.Conditions;
using FlaUI.Core.AutomationElements;
using System.Runtime.InteropServices;
using System.Threading;
using System.CodeDom;
using Windows.Devices.HumanInterfaceDevice;

namespace DesktopDSPTTest // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        static void funct2()
        {
            var app = Application.Launch("path_to_your_application");
            var automation = new UIA3Automation();
            var mainWindow = app.GetMainWindow(automation);

            // Assuming you have a list of files
            //foreach (var file in listOfFiles)
            {
                //  if (file.Name.Contains("DOC") || file.Name.Contains("PPT") || file.Name.Contains("XLS") || file.Name.Contains("TXT"))
                {
                    //file.AsButton().Invoke();
                    var window = new UIA3Automation();
                    var desktopWindow = window.GetDesktop();
                    //desktopWindow = WaitForElement(() => desktopWindow.FindFirstChild(cr => cr.ByName("// Here want to put name that start with... or contains...")));
                    //var app = FlaUI.Core.Application.Attach(desktopWindow.Properties.ProcessId);
                    //var application = app.GetMainWindow(new UIA3Automation());
                    //CloseingProcess(application.Name);
                }
            }
        }
        static void funct3()
        {
            // not working
            var app = Application.Launch(@"C:\Program Files (x86)\CPSSoft\ACCURATE4 Enterprise\Accurate.exe");
            var automation1 = new UIA3Automation();
            var DWindow = automation1.GetDesktop();

            var automation2 = new UIA2Automation();
            var mainWindow = app.GetMainWindow(automation2);
            // Wait until Accurate window ready
            FlaUI.Core.Input.Wait.UntilResponsive(mainWindow.FindFirstChild(), TimeSpan.FromMilliseconds(5000));
            DWindow = WaitForElement(() => DWindow.FindFirstChild(cr => cr.ByName("Welcome to Accurate")));
            if (DWindow != null)
            {
                var xx = DWindow.Properties.GetType();
                DWindow.Click();
                //var appw = FlaUI.Core.Application.Attach(DWindow.Properties.ProcessId);
                //app.Close();
            }



        }
        static void closeWarningDialogBox(string Title)
        {
            var hWnd = FindWindow(null, Title);
            if (hWnd != IntPtr.Zero)
            {
                SendMessage(hWnd, WM_CLOSE, 0, 0);
                Console.WriteLine("Closing dialog message box.");
            }
        }

        const UInt32 WM_CLOSE = 0x0010;

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        static Application appx;
        static Window DesktopWindow;
        static UIA3Automation automationUIA3 = new UIA3Automation();
        static ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());
        static AutomationElement window = automationUIA3.GetDesktop();
        static int step = 0;
        static string LoginId = "supervisor";
        static string LoginPassword = "supervisor";

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

        static void Main(string[] args)
        {
            try
            {
                if (!OpenAppAndDB())
                {
                    Console.WriteLine("Application Automation failed !!");
                    return;
                }
                if (!Login())
                {
                    Console.WriteLine("Application Automation failed !!");
                    return;
                }
                Thread.Sleep(60000);

                /* Try to navigare and open 'Sales' report */
                if (!OpenReport01("sales"))
                {
                    Console.WriteLine("Application Automation failed !!");
                    return;
                }
                /* Download opened report on screen */
                if (!DownloadReport("sales"))
                {
                    Console.WriteLine("Application Automation failed !!");
                    return;
                }
                /* Closing Workspaces that contain all report tab */
                if (!ClosingWorkspace())
                {
                    Console.WriteLine("Application Automation failed !!");
                    return;
                }
                /* Try to navigare and open 'Sales' report */
                if (!OpenReport01("ar"))
                {
                    Console.WriteLine("Application Automation failed !!");
                    return;
                }
                /* Download opened report on screen */
                if (!DownloadReport("ar"))
                {
                    Console.WriteLine("Application Automation failed !!");
                    return;
                }
                /* Closing Workspaces that contain all report tab */
                if (!ClosingWorkspace())
                {
                    Console.WriteLine("Application Automation failed !!");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error => {ex.ToString()}");
            }
            finally
            {
                if (automationUIA3 != null)
                { automationUIA3 = null; }
                {

            }
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
                    Console.WriteLine($"[Step #1 Quitting, end of ClosingWorkspace automation function !!");
                    return false;
                }
                eleMain.SetForeground();
                Thread.Sleep(1000);

                var ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                if (ele is null)
                {
                    Console.WriteLine($"[Step #2 Quitting, end of OpenApp automation function !!");
                    return false;
                }
                Console.WriteLine(ele.Properties.ClassName.ToString());
                ele.SetForeground();
                Console.WriteLine(ele.Properties.ClassName.ToString());
                ele.Focus();
                Thread.Sleep(1000);

                System.Windows.Forms.SendKeys.SendWait("%o");
                Console.WriteLine("Sending keys 'ALT+o'...");
                Thread.Sleep(1000);

                System.Windows.Forms.SendKeys.SendWait("A");
                Console.WriteLine("Then sending key 'a'...");
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
                Console.WriteLine("Closing 'Welcome to Accurate' window.");

                Thread.Sleep(1500);

                /* Closing Warning dialog box */
                var ele3 = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Warning")));
                if (ele3 is null)
                {
                    Console.WriteLine($"[Step #{step += 1}] Quitting, end of OpenApp automation function !!");
                    return false;
                }
                ele3.FindFirstChild(cf => cf.ByName("OK")).AsButton().Click();

                var ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                if (ele is null)
                {
                    Console.WriteLine($"[Step #{step += 1}] Quitting, end of OpenApp automation function !!");
                    return false;
                }
                Console.WriteLine(ele.Properties.ClassName.ToString());
                ele.SetForeground();

                ele = WaitForElement(() => ele.FindFirstChild(cr => cr.ByName("File")));
                if (ele is null)
                {
                    Console.WriteLine($"[Step #{step += 1}] Quitting, end of OpenApp automation function !!");
                    return false;
                }
                Console.WriteLine(ele.Properties.Name.ToString());
                ele.Click();

                //ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByAutomationId("Item 3")));
                //if (ele is null)
                //{
                //    Console.WriteLine("[Step #3] Quitting, end of OpenApp automation function !!");
                //    return false;
                //}
                //ele.AsMenuItem().Click();
                System.Windows.Forms.SendKeys.SendWait("%F");
                Console.WriteLine("Sending keys 'ALT+F'...");
                Thread.Sleep(1000);

                //ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Open Database")));
                //if (ele is null)
                //{
                //    Console.WriteLine("[Step #4] Quitting, end of OpenApp automation function !!");
                //    return false;
                //}
                System.Windows.Forms.SendKeys.SendWait("o");
                Console.WriteLine("Then sending key 'o'...");
                Thread.Sleep(1000);

                //Open Database
                ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Open Database")));
                if (ele is null)
                {
                    Console.WriteLine($"[Step #{step += 1}] Quitting, end of OpenApp automation function !!");
                    return false;
                }
                ele.Focus();

                var ele2 = WaitForElement(() => ele.FindFirstChild(cr => cr.ByClassName("TEdit")));
                if (ele2 is null)
                {
                    Console.WriteLine($"[Step #{step += 1}] Quitting, end of OpenApp automation function !!");
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
                    Console.WriteLine($"[Step #{step += 1}] Quitting, end of OpenApp automation function !!");
                    return false;
                }
                ele.Click();

                return true;
            }
            catch
            {
                Console.WriteLine("Quitting, end of OpenApp automation function !!");
                return false;
            }
        }

        static bool Login()
        {
            try
            {
                var ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Login")));
                if (ele is null)
                {
                    Console.WriteLine($"[Step #{step += 1}] Quitting, end of login automation function !!");
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
                Console.WriteLine($"[Step #{step += 1}] Quitting, end of login automation function !!");
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
                    Console.WriteLine($"[Step #1 Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                Console.WriteLine("Class name is {0}", ele1.Properties.ClassName.ToString());
                ele1.SetForeground();
                Thread.Sleep(500);

                //AutomationId	3540388
                var ele = ele1.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar"));
                if (ele is null)
                {
                    Console.WriteLine($"[Step #2] Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                Console.WriteLine(ele.Properties.ClassName.ToString());
                //ele.Focus();

                System.Windows.Forms.SendKeys.SendWait("%R");
                Console.WriteLine("Sending keys 'ALT+R'...");
                Thread.Sleep(1000);

                System.Windows.Forms.SendKeys.SendWait("i");
                Console.WriteLine("Then sending key 'I'...");
                Thread.Sleep(1000);

                /* Travesing to 'Index to Reports' */
                //window = automationUIA3.GetDesktop();
                var eleParent = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Index to Reports")));
                if (eleParent is null)
                {
                    Console.WriteLine($"[Step #3] Quitting, end of OpenReport0\\{rptType}1 automation function !!");
                    return false;
                }
                Console.WriteLine("Class name is {0}", eleParent.Properties.ClassName.ToString());
                Thread.Sleep(1000);

                if (rptType == "sales")
                {
                    //Sales Reports
                    ele = eleParent.FindFirstDescendant(cf => cf.ByName("Sales Reports"));
                    if (ele is null)
                    {
                        Console.WriteLine($"[Step #4] Quitting, end of OpenReport01\\{rptType} automation function !!");
                        return false;
                    }
                    ele.Click();
                    Thread.Sleep(500);

                    //Sales By Customer Detail
                    ele = eleParent.FindFirstDescendant(cf => cf.ByName("Sales By Customer Detail"));
                    if (ele is null)
                    {
                        Console.WriteLine($"[Step #5] Quitting, end of OpenReport01\\{rptType} automation function !!"); ;
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
                        Console.WriteLine($"[Step #9] Quitting, end of OpenReport01\\{ rptType } automation function !!");
                        return false;
                    }
                    ele.Click();
                    Thread.Sleep(500);

                    //Invoices Paid Summary
                    ele = eleParent.FindFirstDescendant(cf => cf.ByName("Invoices Paid Summary"));
                    if (ele is null)
                    {
                        Console.WriteLine($"[Step #10] Quitting, end of OpenReport01\\{rptType} automation function !!"); ;
                        return false;
                    }
                    ele.DoubleClick();
                    Thread.Sleep(3000);
                }

                //Report Format
                eleParent = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Report Format")).AsWindow());
                if (eleParent is null)
                {
                    Console.WriteLine($"[Step #6] Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                Console.WriteLine(eleParent.Properties.ClassName.ToString());
                eleParent.Focus();
                Thread.Sleep(2000);

                //  Filters && Parameters => Under 'Desktop' windows
                ele = eleParent.FindFirstDescendant(cf => cf.ByName("Filters && Parameters", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                //ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Filters && Parameters", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring)));
                if (ele is null)
                {
                    Console.WriteLine($"[Step #7] Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                Console.WriteLine(ele.Properties.ClassName.ToString());
                ele.Focus();
                Thread.Sleep(500);

                //TabDateFromTo
                ele = ele.FindFirstDescendant(cr => cf.ByName("TabDateFromTo"));
                if (ele is null)
                {
                    Console.WriteLine($"[Step #8] Quitting, end of OpenReport01\\{rptType} automation function !!");
                    return false;
                }
                ele.Focus();

                /* Sending Report Date Parameters */
                AutomationElement[] ArrEle = ele.FindAllDescendants(cf => cf.ByClassName("TDateEdit"));
                if (ArrEle.Length > 0)
                {
                    //TDateEdit
                    Console.WriteLine($"Number of DATE parameter on screen is : {ArrEle.Length}");
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
                Console.WriteLine($"[Exception] Quitting, end of OpenReport01\\{rptType} automation function !!");
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
                    Console.WriteLine($"[Step #1 Quitting, end of DownloadReport automation function !!");
                    return false;
                }
                Console.WriteLine("Class name is {0}", ele1.Properties.ClassName.ToString());
                Thread.Sleep(500);

                var ele = ele1.FindFirstDescendant(cf => cf.ByName("PriviewToolBar"));
                if (ele1 is null)
                {
                    Console.WriteLine($"[Step #2 Quiting, end of DownloadReport automation function !!");
                    return false;
                }
                Console.WriteLine("Class name is {0}", ele.Properties.ClassName.ToString());

                //Export settings
                ele.FindFirstChild(cf.ByName(("Export"))).AsButton().Click();
                Thread.Sleep(1000);


                /* The export button action resulting new window opened */
                ele1 = window.FindFirstDescendant(cf => cf.ByName("Export to Excel"));
                if (ele1 is null)
                {
                    Console.WriteLine($"[Step #3 Quitting, end of DownloadReport automation function !!");
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
            catch (Exception ex) { Console.WriteLine(ex.ToString()); return false; }
        }

        private static bool SavingFileDialog(string reportName)
        {
            ///window = automationUIA3a.GetDesktop();
            //Save As
            //var ele1 = window.FindFirstDescendant(cf.ByClassName("#32770"));
            //if (ele1 is null)
            //{
            //    Console.WriteLine($"[Step #1] Quitting, end of OpenReport\\SavingFileDialog automation function !!");
            //    return false;
            //}
            //ele1.Focus();

            //Edit
            System.Windows.Forms.SendKeys.SendWait("%n");
            Console.WriteLine("Sending keys 'ALT+n'...");
            Thread.Sleep(500);

            /* Delete all excel files in temp folder */
                DeleteExcelFiles(@"C:\temp");

            System.Windows.Forms.SendKeys.SendWait($@"C:\temp\{reportName}.xls");
            Thread.Sleep(500);

            //Save
            var ele = window.FindFirstDescendant(cf => cf.ByName("Save"));
            if (ele is null)
            {
                Console.WriteLine($"[Step #2] Quitting, end of OpenReport\\SavingFileDialog automation function !!");
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
                Console.WriteLine($"[Step #1] Quitting, end of SendingDate automation function !!");
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
                Console.WriteLine($"[Step #2] Quitting, end of OpenReport01 automation function !!");
                return false;
            }

            childEle.Click();
            Thread.Sleep(500);
            childEle.Click();

            Console.WriteLine($"Sending date parameter");

            return true;
        }

        private static void DeleteExcelFiles(string FolderName)
        {
            // Delete Excel files
            var supportFiles = Directory.EnumerateFiles(FolderName, "*.xl*");
            foreach (var excelFile in supportFiles)
            {
                File.Delete(excelFile);
                Console.WriteLine($"Deleted Excel file: {excelFile}");
            }
        }
    }

}
