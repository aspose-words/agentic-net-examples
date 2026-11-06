using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        var doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
        }

        // Add a simple VBA module.
        string moduleSource = "Sub HelloWorld()\n    MsgBox \"Hello from VBA!\"\nEnd Sub";
        var module = new VbaModule
        {
            Name = "HelloModule",
            SourceCode = moduleSource
        };
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro‑enabled file.
        const string fileName = "Sample.docm";
        doc.Save(fileName, SaveFormat.Docm);

        // Reload the document to work with the persisted VBA project.
        var loadedDoc = new Document(fileName);

        // Enumerate references and log them (no COM references are expected in this sample).
        foreach (VbaReference reference in loadedDoc.VbaProject.References)
        {
            Console.WriteLine($"Reference LibId: {reference.LibId}, Type: {reference.Type}");
        }
    }
}
