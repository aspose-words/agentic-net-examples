using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class ImportVbaModule
{
    static void Main()
    {
        // Output path for the generated macro‑enabled document.
        string outputDocPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docm");

        // Create a new blank document (or load an existing .docm if desired).
        Document doc = new Document();

        // Ensure the document has a VBA project; create one if it does not exist.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject { Name = "MyVbaProject" };
        }

        // VBA source code to import. This replaces reading from an external .bas file.
        string vbaSource = @"Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub";

        // Create a new VBA module and set its properties.
        VbaModule module = new VbaModule
        {
            Name = "ImportedModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = vbaSource
        };

        // Add the module to the document's VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro‑enabled file.
        doc.Save(outputDocPath, SaveFormat.Docm);
        Console.WriteLine($"Document saved to: {outputDocPath}");
    }
}
