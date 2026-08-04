using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file paths for the original and modified documents.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string originalPath = Path.Combine(artifactsDir, "Sample.docm");
        string modifiedPath = Path.Combine(artifactsDir, "Sample_Modified.docm");

        // -----------------------------------------------------------------
        // Create a new macro‑enabled document with a single VBA module.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create and assign a new VBA project.
        VbaProject project = new VbaProject();
        project.Name = "MyVbaProject";
        doc.VbaProject = project;

        // Create a procedural VBA module with simple macro code.
        VbaModule module = new VbaModule();
        module.Name = "Module1";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub";

        // Add the module to the project and save the document.
        doc.VbaProject.Modules.Add(module);
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // Load the saved document and validate the presence of at least one module.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        bool hasValidModule = loadedDoc.HasMacros &&
                              loadedDoc.VbaProject != null &&
                              loadedDoc.VbaProject.Modules != null &&
                              loadedDoc.VbaProject.Modules.Count > 0;

        Console.WriteLine($"Document has macros: {loadedDoc.HasMacros}");
        Console.WriteLine($"VBA project present: {loadedDoc.VbaProject != null}");
        Console.WriteLine($"Modules count: {(loadedDoc.VbaProject?.Modules?.Count ?? 0)}");
        Console.WriteLine($"Ready for modification: {hasValidModule}");

        if (hasValidModule)
        {
            // Retrieve the first module safely.
            VbaModule firstModule = loadedDoc.VbaProject.Modules[0];

            // Guard against null source code.
            string currentSource = firstModule.SourceCode ?? string.Empty;

            // Append a comment to indicate the macro was modified.
            firstModule.SourceCode = currentSource + "\n' Modified by Aspose.Words";

            // Save the modified document.
            loadedDoc.Save(modifiedPath);
            Console.WriteLine($"Modified document saved to: {modifiedPath}");
        }
        else
        {
            Console.WriteLine("No VBA modules found; modification skipped.");
        }
    }
}
