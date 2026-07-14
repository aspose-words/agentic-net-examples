using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define a folder to store the temporary files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the macro‑enabled document.
        string macroDocPath = Path.Combine(outputDir, "Sample.docm");

        // -----------------------------------------------------------------
        // Create a new blank document and add a VBA project with one module.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create a new VBA project and assign a name.
        VbaProject project = new VbaProject();
        project.Name = "Aspose.SampleProject";
        doc.VbaProject = project;

        // Create a procedural VBA module with simple macro code.
        VbaModule module = new VbaModule();
        module.Name = "HelloModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = "Sub Hello()\n    MsgBox \"Hello from VBA!\"\nEnd Sub";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document in a macro‑enabled format.
        doc.Save(macroDocPath);

        // ---------------------------------------------------------------
        // Load the saved document and enumerate all VBA modules it contains.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(macroDocPath);

        // Ensure the document actually has a VBA project.
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            VbaModuleCollection modules = loadedDoc.VbaProject.Modules;
            Console.WriteLine($"VBA Project Name: {loadedDoc.VbaProject.Name}");
            Console.WriteLine($"Number of modules: {modules.Count}");

            // Iterate through each module and output its name and source code.
            foreach (VbaModule mod in modules)
            {
                // Guard against null source code.
                string source = mod.SourceCode ?? string.Empty;
                Console.WriteLine($"Module Name: {mod.Name}");
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
