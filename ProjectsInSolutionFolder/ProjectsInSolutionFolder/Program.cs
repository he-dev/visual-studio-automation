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

            var projectPaths = new[]
            {
                @"C:\Home\Projects\Reusable\Reusable.Core\Reusable.Core.csproj",
                @"C:\Home\Projects\Reusable\Reusable.Flexo\Reusable.Flexo.csproj",
                @"C:\Home\Projects\Reusable\Reusable.IOnymous\Reusable.IOnymous.csproj",
            };

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
