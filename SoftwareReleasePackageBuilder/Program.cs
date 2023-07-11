using EnvDTE;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using TCatSysManagerLib;
using System.Xml;
using System.Reflection;

namespace SoftwareReleasePackageBuilder
{
    class Program
    {
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll")]
        static extern int GetRunningObjectTable(uint reserved, out System.Runtime.InteropServices.ComTypes.IRunningObjectTable pprot);

        static void Main(string[] args)
        {
            Console.Title = Assembly.GetExecutingAssembly().GetName().Name + "     " + Assembly.GetExecutingAssembly().GetName().Version;

            List<string> ProjectVariants = new List<string>();

            Console.WriteLine("Building solution variants...");

            EnvDTE.DTE dte;
            ITcSysManager15 sysMan;

            try
            {
                dte = attachToExistingDte();
                Solution solution = dte.Solution;
                Project project = solution.Projects.Item(1);
                sysMan = project.Object;
            }
            catch (Exception err)
            {
                Console.WriteLine("Error connecting to development environment!\n" + err.Message + '\n' + err.StackTrace);
                Console.WriteLine("\n\nPress any key to exit");
                Console.ReadKey();
                return;
            }
            
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(sysMan.ProjectVariantConfig);
            XmlNodeList elemList = xmlDocument.GetElementsByTagName("Name");

            for (int i = 0; i < elemList.Count; i++)
            {
                ProjectVariants.Add(elemList[i].InnerText);
            }

            bool IsRebuilt = false; //rebuild is necessary to save the correct build date and time, but only needs to happen for the first variant
            foreach (string variant in ProjectVariants)
            {
                bool BuildFinished = false;

                sysMan.CurrentProjectVariant = variant;
                Console.Write(variant);

                if (!IsRebuilt)
                {
                    dte.Solution.SolutionBuild.Clean(true);
                    IsRebuilt = true;
                }

                System.Threading.Thread thread = new System.Threading.Thread(() =>
                {
                    dte.Solution.SolutionBuild.Build(true);
                    BuildFinished = true;
                });
                thread.Start();

                while (!BuildFinished)
                {
                    Console.Write(".");
                    System.Threading.Thread.Sleep(500);
                }
                Console.WriteLine();
            }

            try
            {
                FileInfo fileInfo = new FileInfo(dte.Solution.FullName);
                DirectoryInfo BootDir = new DirectoryInfo(fileInfo.Directory.FullName + @"\APMACS_1\_Boot\");
                if (Directory.Exists(BootDir.FullName + @"\TC"))
                {
                    Directory.Delete(BootDir.FullName + @"\TC", true);
                }               
                BootDir.CreateSubdirectory("TC");
                Directory.CreateDirectory(BootDir.FullName + @"\TC\CurrentConfigFiles\");
                Directory.Move(BootDir.FullName + @"TwinCAT RT (x64)\Repository", BootDir.FullName + @"\TC\Repository");

                foreach (string variant in ProjectVariants)
                {
                    if (File.Exists(BootDir.FullName + variant + @"\CurrentConfig.xml"))
                    {
                        File.Delete(BootDir.FullName + variant + @"\CurrentConfig.xml");
                    }
                    File.Move(BootDir.FullName + variant + @"\TwinCAT RT (x64)\CurrentConfig.xml", BootDir.FullName + variant + @"\CurrentConfig.xml");
                    Directory.Delete(BootDir.FullName + variant + @"\TwinCAT RT (x64)\", true);
                    Directory.Move(BootDir.FullName + variant, BootDir.FullName + @"\TC\CurrentConfigFiles\" + variant);
                }

                //copy this file for commissioning
                File.Copy(BootDir.FullName + @"\TC\CurrentConfigFiles\FA_1\CurrentConfig.xml", BootDir.FullName + @"\TC\CurrentConfig.xml");

                string fileText = "The current config file in this folder is the default used for commissioning.";
                File.WriteAllText(BootDir.FullName + @"\TC\" + fileText + "txt", fileText);
            }
            catch (Exception err)
            {
                Console.WriteLine("Error while doing file operations!\n" + err.Message + '\n' + err.StackTrace);
                Console.WriteLine("\n\nPress any key to exit");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("Finished");
            System.Threading.Thread.Sleep(500);
        }

        static public EnvDTE.DTE attachToExistingDte()
        {
            string progId = "";
            EnvDTE.DTE dte = null;
            Hashtable dteInstances = GetIDEInstances(false, progId);
            IDictionaryEnumerator hashtableEnumerator = dteInstances.GetEnumerator();

            while (hashtableEnumerator.MoveNext())
            {
                EnvDTE.DTE dteTemp = hashtableEnumerator.Value as EnvDTE.DTE;
                if (dteTemp.Solution.FullName.Contains(@"APMACS_1\APMACS_1.sln"))
                {
                    dte = dteTemp;
                }
            }
            return dte;
        }

        public static Hashtable GetIDEInstances(bool openSolutionsOnly, string progId)
        {
            Hashtable runningIDEInstances = new Hashtable();
            Hashtable runningObjects = GetRunningObjectTable();
            IDictionaryEnumerator rotEnumerator = runningObjects.GetEnumerator();
            while (rotEnumerator.MoveNext())
            {
                string candidateName = (string)rotEnumerator.Key;
                if (!candidateName.StartsWith("!" + progId))
                    continue;
                EnvDTE.DTE ide = rotEnumerator.Value as EnvDTE.DTE;
                if (ide == null)
                    continue;
                if (openSolutionsOnly)
                {
                    try
                    {
                        string solutionFile = ide.Solution.FullName;
                        if (solutionFile != String.Empty)
                            runningIDEInstances[candidateName] = ide;
                    }
                    catch { }
                }
                else
                    runningIDEInstances[candidateName] = ide;
            }
            return runningIDEInstances;
        }

        public static Hashtable GetRunningObjectTable()
        {
            Hashtable result = new Hashtable();

            IntPtr numFetched = new IntPtr();
            IRunningObjectTable runningObjectTable;
            IEnumMoniker monikerEnumerator;
            IMoniker[] monikers = new IMoniker[1];

            GetRunningObjectTable(0, out runningObjectTable);
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                IBindCtx ctx;
                CreateBindCtx(0, out ctx);

                string runningObjectName;
                monikers[0].GetDisplayName(ctx, null, out runningObjectName);

                object runningObjectVal;
                runningObjectTable.GetObject(monikers[0], out runningObjectVal);

                result[runningObjectName] = runningObjectVal;
            }

            return result;
        }
    }
}
