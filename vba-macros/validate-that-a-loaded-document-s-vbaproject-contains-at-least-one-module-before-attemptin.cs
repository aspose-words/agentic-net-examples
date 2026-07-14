using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docm");
        string modifiedPath = Path.Combine(Directory.GetCurrentDirectory(), "Modified.docm");

        // -----------------------------------------------------------------
        // Step 1: Create a blank document and add a VBA project with one module.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create a new VBA project.
        VbaProject vbaProject = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = vbaProject;

        // Create a procedural module with simple VBA code.
        VbaModule module = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub"
        };
        // Add the module to the project.
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro-enabled file.
        doc.Save(originalPath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // Step 2: Load the saved document and validate the presence of at least one module.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        bool hasAtLeastOneModule = loadedDoc.HasMacros &&
                                   loadedDoc.VbaProject != null &&
                                   loadedDoc.VbaProject.Modules.Count > 0;

        if (hasAtLeastOneModule)
        {
            // -----------------------------------------------------------------
            // Step 3: Perform a modification – prepend a comment to the first module's source code.
            // -----------------------------------------------------------------
            VbaModule firstModule = loadedDoc.VbaProject.Modules[0];
            string originalCode = firstModule.SourceCode ?? string.Empty;
            string modifiedCode = "' Modified by Aspose.Words\n" + originalCode;
            firstModule.SourceCode = modifiedCode;

            // Save the modified document.
            loadedDoc.Save(modifiedPath, SaveFormat.Docm);
            Console.WriteLine("Document contained a module. Modification applied and saved to: " + modifiedPath);
        }
        else
        {
            Console.WriteLine("Document does not contain any VBA modules. No modifications performed.");
        }
    }
}
