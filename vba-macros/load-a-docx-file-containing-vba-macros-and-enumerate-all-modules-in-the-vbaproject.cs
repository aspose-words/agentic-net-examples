using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
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
            string moduleCode = @"Sub Hello()
    MsgBox ""Hello from VBA!""
End Sub";

            VbaModule module = new VbaModule
            {
                Name = "TestModule",
                SourceCode = moduleCode
            };

            doc.VbaProject.Modules.Add(module);
        }

        // Save the document as a macro‑enabled file.
        const string fileName = "SampleMacroDocument.docm";
        doc.Save(fileName, SaveFormat.Docm);

        // Load the saved document.
        Document loadedDoc = new Document(fileName);

        // Enumerate all VBA modules in the loaded document.
        if (loadedDoc.VbaProject != null)
        {
            foreach (VbaModule vbaModule in loadedDoc.VbaProject.Modules)
            {
                string source = vbaModule.SourceCode ?? string.Empty;
                Console.WriteLine($"Module Name: {vbaModule.Name}");
                Console.WriteLine("Source Code:");
                Console.WriteLine(source);
                Console.WriteLine(new string('-', 40));
            }
        }
        else
        {
            Console.WriteLine("No VBA project found in the document.");
        }
    }
}
