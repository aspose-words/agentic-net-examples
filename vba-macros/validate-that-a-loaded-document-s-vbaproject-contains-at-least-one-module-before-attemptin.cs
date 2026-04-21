using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Path for the sample macro-enabled document.
        const string samplePath = "Sample.docm";
        const string modifiedPath = "Modified.docm";

        // -----------------------------------------------------------------
        // Step 1: Create a blank document and ensure it has a VBA project.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // If the document does not already contain a VBA project, create one.
        if (doc.VbaProject == null)
        {
            VbaProject project = new VbaProject();
            project.Name = "SampleProject";
            doc.VbaProject = project;
        }

        // Create a simple VBA module.
        VbaModule module = new VbaModule();
        module.Name = "SampleModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub";

        // Add the module to the project's module collection.
        doc.VbaProject.Modules.Add(module);

        // Save the document in a macro-enabled format.
        doc.Save(samplePath);

        // -----------------------------------------------------------------
        // Step 2: Load the document and validate the presence of at least one module.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(samplePath);

        // Validate that the document actually contains a VBA project.
        if (loadedDoc.VbaProject == null || loadedDoc.VbaProject.Modules.Count == 0)
        {
            Console.WriteLine("The loaded document does not contain any VBA modules. No modifications will be performed.");
            return;
        }

        Console.WriteLine($"Document contains {loadedDoc.VbaProject.Modules.Count} VBA module(s). Proceeding with modification.");

        // -----------------------------------------------------------------
        // Step 3: Modify the source code of the first module safely.
        // -----------------------------------------------------------------
        VbaModule firstModule = loadedDoc.VbaProject.Modules[0];

        // Guard against null source code.
        string originalSource = firstModule.SourceCode ?? string.Empty;

        // Append a comment line to the existing source code.
        string modifiedSource = originalSource + "\n' Added by Aspose.Words example";

        firstModule.SourceCode = modifiedSource;

        // Save the modified document.
        loadedDoc.Save(modifiedPath);

        Console.WriteLine($"Modification completed. Saved as '{modifiedPath}'.");
    }
}
