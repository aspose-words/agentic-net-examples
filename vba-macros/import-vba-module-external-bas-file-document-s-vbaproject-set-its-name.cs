using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class ImportVbaModule
{
    static void Main()
    {
        // Paths to the source .bas file, the input document (optional), and the output document.
        string basFilePath = @"C:\Path\To\module.bas";
        string outputDocPath = @"C:\Path\To\output.docm";

        // Create a new blank document. If you already have a macro‑enabled document, load it instead:
        // Document doc = new Document(@"C:\Path\To\input.docm");
        Document doc = new Document();

        // Ensure the document has a VBA project; create one if it does not exist.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
            doc.VbaProject.Name = "MyVbaProject";
        }

        // Read the VBA source code from the external .bas file.
        string vbaSource = File.ReadAllText(basFilePath);

        // Create a new VBA module, set its name, type, and source code.
        VbaModule module = new VbaModule
        {
            Name = "ImportedModule",               // Desired module name
            Type = VbaModuleType.ProceduralModule, // Most .bas files are procedural modules
            SourceCode = vbaSource
        };

        // Add the module to the document's VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro‑enabled file.
        doc.Save(outputDocPath, SaveFormat.Docm);
    }
}
