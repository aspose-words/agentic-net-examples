using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string macroDocPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleWithMacros.docm");

        // -----------------------------------------------------------------
        // Create a blank document and add a VBA project with a simple module.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = vbaProject;

        // Create a procedural VBA module with some sample code.
        VbaModule vbaModule = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub"
        };

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Save the document in macro-enabled format.
        doc.Save(macroDocPath);

        // ---------------------------------------------------------------
        // Load the saved document and enumerate all VBA modules in it.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(macroDocPath);

        // Ensure the document actually contains a VBA project.
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            VbaModuleCollection modules = loadedDoc.VbaProject.Modules;
            Console.WriteLine($"VBA Project Name: {loadedDoc.VbaProject.Name}");
            Console.WriteLine($"Modules Count: {modules.Count}");

            foreach (VbaModule module in modules)
            {
                // Guard against null source code.
                string source = module.SourceCode ?? string.Empty;
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
