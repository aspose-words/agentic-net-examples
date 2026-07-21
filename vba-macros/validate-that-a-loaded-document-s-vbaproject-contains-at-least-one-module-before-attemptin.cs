using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the initial macro-enabled document.
        string originalPath = Path.Combine(outputDir, "sample.docm");
        // Path for the document after modification.
        string modifiedPath = Path.Combine(outputDir, "modified.docm");

        // -------------------------------------------------
        // 1. Create a blank document and add a VBA project.
        // -------------------------------------------------
        Document doc = new Document();

        // Create a new VBA project.
        VbaProject project = new VbaProject
        {
            Name = "Aspose.Project"
        };

        // Create a simple procedural module with some VBA code.
        VbaModule module = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Hello()\n    MsgBox \"Hello from VBA!\"\nEnd Sub"
        };

        // Add the module to the project and assign the project to the document.
        project.Modules.Add(module);
        doc.VbaProject = project;

        // Save the document in a macro‑enabled format.
        doc.Save(originalPath);

        // -------------------------------------------------
        // 2. Load the saved document and validate VBA modules.
        // -------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Ensure the document actually contains a VBA project and at least one module.
        bool hasModules = loadedDoc.HasMacros &&
                          loadedDoc.VbaProject != null &&
                          loadedDoc.VbaProject.Modules != null &&
                          loadedDoc.VbaProject.Modules.Count > 0;

        if (hasModules)
        {
            // Access the first module safely.
            VbaModule firstModule = loadedDoc.VbaProject.Modules[0];

            // Guard against null source code.
            string currentSource = firstModule.SourceCode ?? string.Empty;

            // Append a comment to demonstrate modification.
            string updatedSource = currentSource + "\n' Added by Aspose.Words example";

            firstModule.SourceCode = updatedSource;

            // Save the modified document.
            loadedDoc.Save(modifiedPath);

            Console.WriteLine("Document contained VBA modules. First module was updated and saved to:");
            Console.WriteLine(modifiedPath);
        }
        else
        {
            Console.WriteLine("The loaded document does not contain any VBA modules. No modifications were made.");
        }
    }
}
