using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Vba;

public class ExportVbaModules
{
    public static void Main()
    {
        // Prepare paths for the temporary macro-enabled document and the output ZIP archive.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docm");
        string zipPath = Path.Combine(Directory.GetCurrentDirectory(), "VbaModules.zip");

        // -----------------------------------------------------------------
        // 1. Create a sample macro-enabled document with a few VBA modules.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Ensure the document has a VBA project.
        VbaProject project = new VbaProject { Name = "SampleProject" };
        doc.VbaProject = project;

        // Helper to add a module.
        void AddModule(string name, string code, VbaModuleType type = VbaModuleType.ProceduralModule)
        {
            VbaModule module = new VbaModule
            {
                Name = name,
                Type = type,
                SourceCode = code
            };
            doc.VbaProject.Modules.Add(module);
        }

        AddModule("ModuleOne", "Sub HelloWorld()\n    MsgBox \"Hello from ModuleOne\"\nEnd Sub");
        AddModule("ModuleTwo", "Function AddNumbers(a As Integer, b As Integer) As Integer\n    AddNumbers = a + b\nEnd Function");
        AddModule("ClassModule", "Public Sub Greet()\n    MsgBox \"Greetings from ClassModule\"\nEnd Sub", VbaModuleType.ClassModule);

        // Save the document in a macro-enabled format.
        doc.Save(docPath, SaveFormat.Docm);

        // ---------------------------------------------------------------
        // 2. Load the document (demonstrating the load rule) and export modules.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Verify that the document actually contains macros.
        if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
        {
            Console.WriteLine("The document does not contain any VBA macros.");
            return;
        }

        // Create the ZIP archive and add each module as a separate entry.
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
        {
            foreach (VbaModule module in loadedDoc.VbaProject.Modules)
            {
                // Use the module name with .bas extension; adjust for class modules if desired.
                string entryName = $"{module.Name}.bas";

                // Guard against null source code.
                string source = module.SourceCode ?? string.Empty;

                // Create a new entry in the ZIP archive.
                ZipArchiveEntry entry = archive.CreateEntry(entryName);

                // Write the source code into the entry.
                using (StreamWriter writer = new StreamWriter(entry.Open()))
                {
                    writer.Write(source);
                }
            }
        }

        // Indicate completion (no interactive input required).
        Console.WriteLine($"Exported {loadedDoc.VbaProject.Modules.Count} VBA module(s) to \"{zipPath}\".");
    }
}
