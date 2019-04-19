using System;
using System.IO;
using System.Linq;
using EnvDTE;
using EnvDTE80;

namespace ProjectsInSolutionFolder
{
    public static class Program
    {
        private const string VisualStudio2019 = "VisualStudio.DTE.16.0";

        public static void Main(string[] args)
        {
            var dteType = Type.GetTypeFromProgID(VisualStudio2019, true);
            var dte = (DTE)Activator.CreateInstance(dteType);

            var sln = (SolutionClass)dte.Solution;
            sln.Open(@"C:\Home\Projects\RoboNuGet\RoboNuGet.sln");

            var lib = FindSolutionFolder(sln, "lib");

            var basePath = @"C:\Home\Projects\Reusable";
            var projectPaths = new[]
            {
                "Reusable.Commander",
                "Reusable.Core",
                "Reusable.Cryptography",
                "Reusable.Deception",
                "Reusable.Flexo",
                "Reusable.IOnymous",
                "Reusable.IOnymous.Http",
                "Reusable.IOnymous.Http.Mailr",
                "Reusable.IOnymous.Mail",
                "Reusable.IOnymous.Mail.Smtp",
                "Reusable.OmniLog",
                "Reusable.OmniLog.Abstractions",
                "Reusable.OmniLog.ColoredConsoleRx",
                "Reusable.OmniLog.NLogRx",
                "Reusable.OmniLog.SemanticExtensions",
                "Reusable.OmniLog.SemanticMiddleware",
                "Reusable.OneTo1",
                "Reusable.SemanticVersion",
                "Reusable.SmartConfig",
                "Reusable.SmartConfig.SqlServer",
                "Reusable.Utilities.AspNetCore",
                "Reusable.Utilities.JsonNet",
                "Reusable.Utilities.NLog",
                "Reusable.Utilities.SqlClient",
            }.Select(n => Path.Combine(basePath, n, $"{n}.csproj"));

            foreach (var path in projectPaths)
            {
                var name = Path.GetFileNameWithoutExtension(path);
                if (lib.Parent.ProjectItems.Cast<ProjectItem>().Any(pi => pi.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                {
                    Console.WriteLine($"{name} - exists");
                }
                else
                {
                    lib.AddFromFile(path);
                    Console.WriteLine($"{name} - added");                    
                }
            }

            dte.Solution.Close(true);

            Console.WriteLine("Done!");
            Console.ReadKey();
        }

        private static SolutionFolder FindSolutionFolder(SolutionClass sln, string folderName)
        {
            var solutionFolder =
                sln.Projects
                    .OfType<Project>()
                    .FirstOrDefault(p => p.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase))
                    ?.Object as SolutionFolder;

            if (solutionFolder is null)
            {
                solutionFolder = (sln as Solution2).AddSolutionFolder(folderName).Object as SolutionFolder;
            }

            return solutionFolder;
        }
    }
}
