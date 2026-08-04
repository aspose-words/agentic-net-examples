using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaMacroEnumeration
{
    public class Program
    {
        public static void Main()
        {
            // Define paths for the temporary macro-enabled document.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);
            string macroDocPath = Path.Combine(outputDir, "SampleMacroDocument.docm");

            // -----------------------------------------------------------------
            // Step 1: Create a blank document and add a VBA project with a module.
            // -----------------------------------------------------------------
            Document doc = new Document();

            // Create a new VBA project and assign a name.
            VbaProject vbaProject = new VbaProject
            {
                Name = "SampleProject"
            };

            // Create a procedural VBA module with simple source code.
            VbaModule vbaModule = new VbaModule
            {
                Name = "SampleModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub"
            };

            // Add the module to the VBA project and assign the project to the document.
            vbaProject.Modules.Add(vbaModule);
            doc.VbaProject = vbaProject;

            // Save the document in a macro-enabled format (.docm).
            doc.Save(macroDocPath);

            // -----------------------------------------------------------------
            // Step 2: Load the saved document and enumerate its VBA modules.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(macroDocPath);

            // Ensure the document actually contains a VBA project.
            if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
            {
                VbaModuleCollection modules = loadedDoc.VbaProject.Modules;
                Console.WriteLine($"VBA Project Name: {loadedDoc.VbaProject.Name}");
                Console.WriteLine($"Number of modules: {modules.Count}");

                // Iterate through each module and output its name and source code.
                foreach (VbaModule module in modules)
                {
                    string source = module.SourceCode ?? string.Empty; // Guard against null source.
                    Console.WriteLine($"Module Name: {module.Name}");
                    Console.WriteLine("Source Code:");
                    Console.WriteLine(source);
                    Console.WriteLine(new string('-', 40));
                }
            }
            else
            {
                Console.WriteLine("The loaded document does not contain any VBA macros.");
            }
        }
    }
}
