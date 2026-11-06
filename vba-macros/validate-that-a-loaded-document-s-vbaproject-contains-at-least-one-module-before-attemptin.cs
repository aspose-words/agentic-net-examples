using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare paths for the sample and modified documents.
        string samplePath = Path.Combine(Path.GetTempPath(), "Sample.docm");
        string modifiedPath = Path.Combine(Path.GetTempPath(), "Modified.docm");

        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
        }

        // Add a simple VBA module if none exist.
        if (doc.VbaProject.Modules.Count == 0)
        {
            VbaModule module = new VbaModule();
            module.Name = "Module1";
            module.SourceCode =
                "Sub Hello()\n" +
                "    MsgBox \"Hello from VBA!\"\n" +
                "End Sub";
            doc.VbaProject.Modules.Add(module);
        }

        // Save the document as a macro‑enabled file.
        doc.Save(samplePath, SaveFormat.Docm);

        // Load the document back.
        Document loadedDoc = new Document(samplePath);

        // Validate that the loaded document contains a VBA project with at least one module.
        if (loadedDoc.VbaProject != null && loadedDoc.VbaProject.Modules.Count > 0)
        {
            // Perform a simple modification: append a comment to the first module's source code.
            VbaModule firstModule = loadedDoc.VbaProject.Modules[0];
            string originalSource = firstModule.SourceCode ?? string.Empty;
            string modifiedSource = originalSource + "\n' Added by C# automation";
            firstModule.SourceCode = modifiedSource;

            // Save the modified document.
            loadedDoc.Save(modifiedPath, SaveFormat.Docm);
            Console.WriteLine($"Document validated and modified. Saved to: {modifiedPath}");
        }
        else
        {
            Console.WriteLine("The loaded document does not contain a VBA project with any modules. No modifications were made.");
        }

        // Clean up temporary files (optional).
        // File.Delete(samplePath);
        // File.Delete(modifiedPath);
    }
}
