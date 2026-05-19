using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file names for the sample and the modified documents.
        string samplePath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docm");
        string modifiedPath = Path.Combine(Directory.GetCurrentDirectory(), "Modified.docm");

        // -------------------------------------------------
        // 1. Create a blank document and add a VBA project.
        // -------------------------------------------------
        Document doc = new Document();

        // Create a new VBA project and give it a name.
        VbaProject project = new VbaProject();
        project.Name = "Aspose.SampleProject";

        // Create a procedural module with simple VBA code.
        VbaModule module = new VbaModule();
        module.Name = "SampleModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub";

        // Add the module to the project and assign the project to the document.
        project.Modules.Add(module);
        doc.VbaProject = project;

        // Save the document in a macro‑enabled format.
        doc.Save(samplePath, SaveFormat.Docm);

        // -------------------------------------------------
        // 2. Load the document and validate VBA modules.
        // -------------------------------------------------
        Document loadedDoc = new Document(samplePath);

        bool hasValidModule = false;

        // Ensure the document actually contains a VBA project.
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            // Check that there is at least one module in the project.
            if (loadedDoc.VbaProject.Modules.Count > 0)
            {
                hasValidModule = true;
                Console.WriteLine("Validation succeeded: VBA project contains at least one module.");
            }
            else
            {
                Console.WriteLine("Validation failed: VBA project has no modules.");
            }
        }
        else
        {
            Console.WriteLine("Validation failed: Document does not contain a VBA project.");
        }

        // -------------------------------------------------
        // 3. Modify the first module only if validation passed.
        // -------------------------------------------------
        if (hasValidModule)
        {
            // Retrieve the first module (index 0).
            VbaModule firstModule = loadedDoc.VbaProject.Modules[0];

            // Guard against null source code.
            string originalCode = firstModule.SourceCode ?? string.Empty;

            // Append a comment to the existing VBA code.
            string modifiedCode = originalCode + "\n' Modified by Aspose.Words";

            firstModule.SourceCode = modifiedCode;

            // Save the modified document.
            loadedDoc.Save(modifiedPath, SaveFormat.Docm);
            Console.WriteLine($"Modified document saved to: {modifiedPath}");
        }
        else
        {
            Console.WriteLine("No modifications were performed due to failed validation.");
        }
    }
}
