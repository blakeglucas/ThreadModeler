using CommandLine;
using Inventor;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ThreadModeler
{
    public class CommandLineOptions
    {
        [Option('i', "input", Required = true, HelpText = "Autodesk Part File (.ipt) to process")]
        public string TargetIPT { get; set; }

        [Option('o', "output", Required = false, Default = null, HelpText = "(Default: <input>_threaded.ipt)")]
        public string OutputFile { get; set; }

        [Option('t', "template", Required = false, Default = "./ISO Template.ipt")]
        public string TemplateFile { get; set; }

        [Option("stl", Required = false, Default = false)]
        public bool OutputSTL { get; set; }

    }
    class Program
    {
        [DllImport("ole32.dll")]
        static extern int CreateBindCtx(
        uint reserved,
        out IBindCtx ppbc);

        [DllImport("ole32.dll")]
        public static extern void GetRunningObjectTable(
            int reserved,
            out IRunningObjectTable prot);

        private static Inventor.Application GetRunningInventorApplication()
        {
            IRunningObjectTable Rot = null;
            GetRunningObjectTable(0, out Rot);

            IEnumMoniker monikerEnumerator = null;
            Rot.EnumRunning(out monikerEnumerator);

            monikerEnumerator.Reset();

            IntPtr pNumFetched = new IntPtr();
            IMoniker[] monikers = new IMoniker[1];

            Process[] processlist = Process.GetProcesses();

            IntPtr inventorWHND = new IntPtr();

            foreach (Process process in processlist)
            {
                if (!String.IsNullOrEmpty(process.MainWindowTitle))
                {
                    if (process.ProcessName == "Inventor")
                    {
                        inventorWHND = process.MainWindowHandle;
                    }
                }
            }

            if (inventorWHND == null)
            {
                return null;
            }

            while (monikerEnumerator.Next(1, monikers, pNumFetched) == 0)
            {
                IBindCtx bindCtx;
                CreateBindCtx(0, out bindCtx);
                if (bindCtx == null)
                    continue;

                string displayName;
                monikers[0].GetDisplayName(bindCtx, null, out displayName);
                System.Console.WriteLine(displayName);

                object ComObject;
                Rot.GetObject(monikers[0], out ComObject);

                if (ComObject == null)
                {
                    continue;
                }

                dynamic invDoc = ComObject;
                if (new IntPtr(invDoc.MainFrameHWND) == inventorWHND)
                {
                    return invDoc;
                }

            }
            return null;
        }

        private static Inventor.Application Initialize()
        {
            Inventor.Application app = null;
            for (; app == null; System.Threading.Thread.Sleep(1000))
            {
                // Find the Inventor.Application object with
                // the correct window handle
                app = GetRunningInventorApplication()
                  as Inventor.Application;
                if (app != null)
                {
                    return app;
                }
            }
            return null;
        }

        private static int ThreadModels(CommandLineOptions options)
        {
            var app = Initialize();
            if (app == null)
            {
                return -1;
            }
            ThreadWorker.Initialize(app);
            Toolkit.Initialize(app);
            string _ThreadTemplatePath = new FileInfo("./ISO Template.ipt").FullName;
            PartDocument template = app.Documents.Open(
                _ThreadTemplatePath, false) as PartDocument;
            string TargetIPTPath = new FileInfo(options.TargetIPT).FullName;
            PartDocument TargetIPT = app.Documents.Open(TargetIPTPath) as PartDocument;

            var threadFeatures = TargetIPT.ComponentDefinition.Features.ThreadFeatures;

            var threads = new List<ThreadFeature>();

            foreach (var feature in threadFeatures)
            {
                var thread = feature as ThreadFeature;

                if (thread.Suppressed)
                {
                    continue;
                }

                threads.Add(thread);
            }

            PlanarSketch templateSketch =
                template.ComponentDefinition.Sketches[1];

            if (!ThreadWorker.ModelizeThreads(
                TargetIPT,
                templateSketch,
                threads,
                0.1))
            {
                System.Console.WriteLine("error haha fuck you");
                return -1;
            }

            string TargetPath = System.IO.Path.GetDirectoryName(TargetIPTPath) + System.IO.Path.DirectorySeparatorChar + System.IO.Path.GetFileNameWithoutExtension(TargetIPTPath) + "_threaded.ipt";
            if (options.OutputFile != null)
            {
                TargetPath = new FileInfo(options.OutputFile).FullName;
            }
            TargetIPT.SaveAs(TargetPath, false);
            if (options.OutputSTL)
            {
                var TargetSTLPath = System.IO.Path.GetDirectoryName(TargetIPTPath) + System.IO.Path.DirectorySeparatorChar + System.IO.Path.GetFileNameWithoutExtension(TargetIPTPath) + "_threaded.stl";
                if (options.OutputFile != null)
                {
                    TargetSTLPath = System.IO.Path.GetFileNameWithoutExtension(new FileInfo(options.OutputFile).FullName) + ".stl";
                }
                TargetIPT.SaveAs(TargetSTLPath, true);
            }
            TargetIPT.Close(true);

            return 0;
        }

        static int Main(string[] args)
        {
            return Parser.Default.ParseArguments<CommandLineOptions>(args).MapResult(
                options => ThreadModels(options),
                _ => 1
            );
        }
    }
}

