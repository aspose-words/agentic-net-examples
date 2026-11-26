using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a macro‑enabled document.
        var doc = new Document();

        // Ensure the document has a VBA project.
        doc.VbaProject = new VbaProject();

        // Add a simple module.
        var module1 = new VbaModule
        {
            Name = "Module1",
            SourceCode = "Sub Hello()\n    MsgBox \"Hello from VBA\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(module1);

        // Save the document in macro‑enabled format.
        const string originalPath = "sample.docm";
        doc.Save(originalPath, SaveFormat.Docm);

        // Load the document back.
        var loadedDoc = new Document(originalPath);

        // Validate that the loaded document contains a VBA project with at least one module.
        bool hasModules = loadedDoc.VbaProject != null && loadedDoc.VbaProject.Modules.Count > 0;

        Console.WriteLine($"Document has VBA project: {loadedDoc.VbaProject != null}");
        Console.WriteLine($"Number of VBA modules: {(loadedDoc.VbaProject?.Modules.Count ?? 0)}");

        if (!hasModules)
        {
            Console.WriteLine("No VBA modules found. Skipping modifications.");
            return;
        }

        // Since a module exists, perform a simple modification: add another module.
        var module2 = new VbaModule
        {
            Name = "Module2",
            SourceCode = "Sub Goodbye()\n    MsgBox \"Goodbye from VBA\"\nEnd Sub"
        };
        loadedDoc.VbaProject.Modules.Add(module2);

        // Save the modified document.
        const string modifiedPath = "modified.docm";
        loadedDoc.Save(modifiedPath, SaveFormat.Docm);

        Console.WriteLine($"Modifications applied and saved to '{modifiedPath}'.");
    }
}
