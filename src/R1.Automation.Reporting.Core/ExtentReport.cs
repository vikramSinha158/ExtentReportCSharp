using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Gherkin.Model;
using System.Reflection;
using TechTalk.SpecFlow;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace R1.Automation.Reporting.Core
{
    public class ExtentReport
    {
        /// <summary>Initializes the report.</summary>
        /// <param name="appFolderName">Name of the application folder.</param>
        /// <returns>Returns ExtentReport Object</returns>
        public static AventStack.ExtentReports.ExtentReports InitReport(string appFolderName)
        {
            var folderName = GetDirName();
            string path;
            if (folderName != null || folderName != "")
            {
                path = Path.Combine(folderName.Substring(0, folderName.LastIndexOf("\\bin")), appFolderName + "\\");
                string folder = DateTime.Now.ToString("dd_MMM_yyyy");
                path = path + folder;
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                path = path + "\\";
                folder = DateTime.Now.ToString("dd_MM_yyyy_hh_mm_ss_tt");
                path = path + folder;
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                path = path + "\\";

                ExtentHtmlReporter htmlReporter = new ExtentHtmlReporter(path);
                htmlReporter.Config.Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Standard;
                AventStack.ExtentReports.ExtentReports extent = new AventStack.ExtentReports.ExtentReports();
                extent.AttachReporter(htmlReporter);
                return extent;
            }
            else
                return null;
        }

        /// <summary>Gets the name of the dir.</summary>
        /// <returns>Return current Directory path</returns>
        private static string GetDirName()
        {
            try
            {
                return Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            }
            catch (Exception ex)
            {
                if (ex is PathTooLongException || ex is System.ArgumentException)
                    return "";
                else
                    return null;
            }
        }

        /// <summary>This method is used for Archive Old Folders</summary>
        /// <param name="appFolderName"></param>
        /// <param name="archiveFolder"></param>
        /// <param name="noOfDays"></param>
        /// <param name="SizeInMB"></param>
        public static void ArchiveOldFolders(string appFolderName, string archiveFolder, string noOfDays, string deleteBeforeArchive, string SizeInMB="5")
        {
            int num = Int32.Parse(noOfDays);
            int FileSize = Int32.Parse(SizeInMB);

            var folderName = GetDirName();
            string path = Path.Combine(folderName.Substring(0, folderName.LastIndexOf("\\bin")), appFolderName + "\\");
            string ArchivePath = Path.Combine(folderName.Substring(0, folderName.LastIndexOf("\\bin")), archiveFolder + "\\");

            string[] subdirectoryEntries = Directory.GetDirectories(path);
            foreach (string subdirectory in subdirectoryEntries)
            {
                DirectoryInfo d = new DirectoryInfo(subdirectory);
                long sizeOfDir = DirectorySize(d, true);
                if (d.CreationTime < DateTime.Now.AddDays(-num) || ((double)sizeOfDir) / (1024*1024) > FileSize)
                {
                    if (File.Exists(ArchivePath + d.Name + ".zip") && deleteBeforeArchive.Equals("Yes"))
                    {
                        File.Delete(ArchivePath + d.Name + ".zip");
                    }

                    try
                    {
                        ZipFile.CreateFromDirectory(path + d.Name, ArchivePath + d.Name + ".zip");
                        d.Delete(true);
                    }catch(Exception e)
                    {

                    }
                }

            }


        }

        

        /// <summary>This method is used to find size of a folder</summary>
        /// <param name="dInfo"></param>
        /// <param name="includeSubDir"></param>
        /// <returns>Size of a folder</returns>
        static long DirectorySize(DirectoryInfo dInfo, bool includeSubDir)
        {
            long totalSize = dInfo.EnumerateFiles()
                         .Sum(file => file.Length);
            if (includeSubDir)
            {
                totalSize += dInfo.EnumerateDirectories()
                         .Sum(dir => DirectorySize(dir, true));
            }

            return totalSize; 
        }

        /// <summary>This method is used for delete archived folders</summary>
        /// <param name="appFolderName"></param>
        /// <param name="noOfDays"></param>
        public static void DeleteArchiveFolder(string appFolderName, string noOfDays)
        {
            int num = Int32.Parse(noOfDays);

            var folderName = GetDirName();
            string path = Path.Combine(folderName.Substring(0, folderName.LastIndexOf("\\bin")), appFolderName + "\\");

            string[] subFileEntries = Directory.GetFiles(path);
            foreach (string subFile in subFileEntries)
            {
                FileInfo d = new FileInfo(subFile);
                if (d.CreationTime < DateTime.Now.AddDays(-num))
                    d.Delete();
            }

        }


        /// <summary>Configurations the steps.</summary>
        /// <param name="scenarioContext">The scenario context.</param>
        /// <returns>Returns result as object</returns>
        public object ConfigSteps(ScenarioContext scenarioContext)
        {
            PropertyInfo pInfo = typeof(ScenarioContext).GetProperty("ScenarioExecutionStatus", BindingFlags.Instance | BindingFlags.Public);
            MethodInfo getter = pInfo.GetGetMethod(nonPublic: true);
            object TestResult = getter.Invoke(scenarioContext, null);
            return TestResult;
        }
        /// <summary>Inserts the steps in report.</summary>
        /// <param name="scenarioContext">The scenario context.</param>
        /// <param name="TestResult">The test result.</param>
        /// <param name="sPath">The screenshot path.</param>
        /// <param name="scenario">The scenario under test</param>
        /// <param name="passScreenShot">If set to true, saves screenshot for pass test case.</param>
        /// <param name="failScreenShot">If set to true, saves screenshot for fail test case. </param>
        public void InsertStepsInReport(ScenarioContext scenarioContext, object TestResult, string sPath, ExtentTest scenario, bool passScreenShot, bool failScreenShot)
        {
            var stepType = scenarioContext.StepContext.StepInfo.StepDefinitionType.ToString().ToLower();
            if (scenarioContext.TestError == null && TestResult.ToString() != "StepDefinitionPending")
            {
                if (passScreenShot)
                    SetExeStatusForPass(scenario, scenarioContext, sPath, stepType);
                else
                    SetExeStatusForPassWithoutScreenshot(scenario, scenarioContext, stepType);
            }
            else if (scenarioContext.TestError != null)
            {
                if (failScreenShot)
                    SetExecutionStatusForFail(scenario, scenarioContext, sPath, stepType);
                else
                    SetExeStatusForFailWithoutScreenshot(scenario, scenarioContext, stepType);
            }
            if (scenarioContext.ScenarioExecutionStatus.ToString() == "StepDefinitionPending")
                SetExecutionStatusForPending(scenario, stepType);
        }

        /// <summary>Inserts the steps in report without screenshot support.</summary>
        /// <param name="scenarioContext">The scenario context.</param>
        /// <param name="TestResult">The test result.</param>
        /// <param name="scenario">The scenario under test</param>
        public void InsertStepsInReport(ScenarioContext scenarioContext, object TestResult, ExtentTest scenario)
        {
            var stepType = scenarioContext.StepContext.StepInfo.StepDefinitionType.ToString().ToLower();
            if (scenarioContext.TestError == null && TestResult.ToString() != "StepDefinitionPending")
            {
                SetExeStatusForPassWithoutScreenshot(scenario, scenarioContext, stepType);
            }
            else if (scenarioContext.TestError != null)
            {
                SetExeStatusForFailWithoutScreenshot(scenario, scenarioContext, stepType);
            }
            if (scenarioContext.ScenarioExecutionStatus.ToString() == "StepDefinitionPending")
                SetExecutionStatusForPending(scenario, stepType);
        }

        /// <summary>Sets the executable status for pass.</summary>
        /// <param name="scenario">The scenario.</param>
        /// <param name="scenarioContext">The scenario context.</param>
        /// <param name="sPath">The screenshot path.</param>
        /// <param name="stepType">Type of the step.</param>
        private void SetExeStatusForPass(ExtentTest scenario, ScenarioContext scenarioContext, string sPath, string stepType)
        {
            if (stepType.Equals("given"))
                scenario.CreateNode<Given>(scenarioContext.StepContext.StepInfo.Text).Info("Find Screen Shot:- " + sPath);
            else if (stepType.Equals("when"))
                scenario.CreateNode<When>(scenarioContext.StepContext.StepInfo.Text).Info("Find Screen Shot:- " + sPath);
            else if (stepType.Equals("then"))
                scenario.CreateNode<Then>(scenarioContext.StepContext.StepInfo.Text).Info("Find Screen Shot:- " + sPath);
            else if (stepType.Equals("and"))
                scenario.CreateNode<And>(scenarioContext.StepContext.StepInfo.Text).Info("Find Screen Shot:- " + sPath);
            else if (stepType.Equals("but"))
                scenario.CreateNode<But>(scenarioContext.StepContext.StepInfo.Text).Info("Find Screen Shot:- " + sPath);
        }

        /// <summary>Sets the executable status for pass without screenshot.</summary>
        /// <param name="scenario">The scenario.</param>
        /// <param name="scenarioContext">The scenario context.</param>
        /// <param name="stepType">Type of the step.</param>
        private void SetExeStatusForPassWithoutScreenshot(ExtentTest scenario, ScenarioContext scenarioContext, string stepType)
        {
            if (stepType.Equals("given"))
                scenario.CreateNode<Given>(scenarioContext.StepContext.StepInfo.Text);
            else if (stepType.Equals("when"))
                scenario.CreateNode<When>(scenarioContext.StepContext.StepInfo.Text);
            else if (stepType.Equals("then"))
                scenario.CreateNode<Then>(scenarioContext.StepContext.StepInfo.Text);
            else if (stepType.Equals("and"))
                scenario.CreateNode<And>(scenarioContext.StepContext.StepInfo.Text);
            else if (stepType.Equals("but"))
                scenario.CreateNode<But>(scenarioContext.StepContext.StepInfo.Text);
        }


        /// <summary>Sets the execution status for fail.</summary>
        /// <param name="scenario">The scenario.</param>
        /// <param name="scenarioContext">The scenario context.</param>
        /// <param name="sPath">The screenshot path.</param>
        /// <param name="stepType">Type of the step.</param>
        private void SetExecutionStatusForFail(ExtentTest scenario, ScenarioContext scenarioContext, string sPath, string stepType)
        {
            if (stepType.Equals("given"))
                scenario.CreateNode<Given>(scenarioContext.StepContext.StepInfo.Text).Fail(scenarioContext.TestError.Message).Info("Find Screen Shot:- " + sPath);
            else if (stepType.Equals("when"))
                scenario.CreateNode<When>(scenarioContext.StepContext.StepInfo.Text).Fail(scenarioContext.TestError.Message).Info("Find Screen Shot:- " + sPath);
            else if (stepType.Equals("then"))
                scenario.CreateNode<Then>(scenarioContext.StepContext.StepInfo.Text).Fail(scenarioContext.TestError.Message).Info("Find Screen Shot:- " + sPath);
            else if (stepType.Equals("and"))
                scenario.CreateNode<And>(scenarioContext.StepContext.StepInfo.Text).Fail(scenarioContext.TestError.Message).Info("Find Screen Shot:- " + sPath);
            else if (stepType.Equals("but"))
                scenario.CreateNode<But>(scenarioContext.StepContext.StepInfo.Text).Fail(scenarioContext.TestError.Message).Info("Find Screen Shot:- " + sPath);

        }

        /// <summary>Sets the executable status for fail without screenshot.</summary>
        /// <param name="scenario">The scenario.</param>
        /// <param name="scenarioContext">The scenario context.</param>
        /// <param name="stepType">Type of the step.</param>
        private void SetExeStatusForFailWithoutScreenshot(ExtentTest scenario, ScenarioContext scenarioContext, string stepType)
        {
            if (stepType.Equals("given"))
                scenario.CreateNode<Given>(scenarioContext.StepContext.StepInfo.Text).Fail(scenarioContext.TestError.Message);
            else if (stepType.Equals("when"))
                scenario.CreateNode<When>(scenarioContext.StepContext.StepInfo.Text).Fail(scenarioContext.TestError.Message);
            else if (stepType.Equals("then"))
                scenario.CreateNode<Then>(scenarioContext.StepContext.StepInfo.Text).Fail(scenarioContext.TestError.Message);
            else if (stepType.Equals("and"))
                scenario.CreateNode<And>(scenarioContext.StepContext.StepInfo.Text).Fail(scenarioContext.TestError.Message);
            else if (stepType.Equals("but"))
                scenario.CreateNode<But>(scenarioContext.StepContext.StepInfo.Text).Fail(scenarioContext.TestError.Message);

        }
        /// <summary>Sets the execution status for pending.</summary>
        /// <param name="scenario">The scenario.</param>
        /// <param name="stepType">Type of the step.</param>
        private void SetExecutionStatusForPending(ExtentTest scenario, string stepType)
        {
            if (stepType == "Given")
                scenario.CreateNode<Given>(ScenarioStepContext.Current.StepInfo.Text).Skip("Step Definition Pending");
            else if (stepType == "When")
                scenario.CreateNode<When>(ScenarioStepContext.Current.StepInfo.Text).Skip("Step Definition Pending");
            else if (stepType == "Then")
                scenario.CreateNode<Then>(ScenarioStepContext.Current.StepInfo.Text).Skip("Step Definition Pending");
            else if (stepType == "And")
                scenario.CreateNode<And>(ScenarioStepContext.Current.StepInfo.Text).Skip("Step Definition Pending");
            else if (stepType == "But")
                scenario.CreateNode<But>(ScenarioStepContext.Current.StepInfo.Text).Skip("Step Definition Pending");
        }
    }
}
