using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Vba;

public class ExportVbaModules
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the sample macro-enabled document.
        string docPath = Path.Combine(outputDir, "Sample.docm");

        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = project;

        // Create first VBA module.
        VbaModule module1 = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Hello()\n    MsgBox \"Hello from Module1\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(module1);

        // Create second VBA module.
        VbaModule module2 = new VbaModule
        {
            Name = "Module2",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Goodbye()\n    MsgBox \"Goodbye from Module2\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(module2);

        // Save the document in macro‑enabled format.
        doc.Save(docPath);

        // Load the saved document (demonstrates a typical workflow).
        Document loadedDoc = new Document(docPath);

        // Verify that the document actually contains macros.
        if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
        {
            Console.WriteLine("No VBA project found in the document.");
            return;
        }

        // Path for the resulting ZIP archive.
        string zipPath = Path.Combine(outputDir, "VbaModules.zip");

        // Create the ZIP archive and add each VBA module as a separate file.
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Update))
        {
            foreach (VbaModule vbaModule in loadedDoc.VbaProject.Modules)
            {
                // Ensure source code is not null.
                string source = vbaModule.SourceCode ?? string.Empty;

                // Use the original module name with a .bas extension.
                string entryName = $"{vbaModule.Name}.bas";

                // Create a new entry in the ZIP archive.
                ZipArchiveEntry entry = archive.CreateEntry(entryName);

                // Write the module source code into the entry.
                using (StreamWriter writer = new StreamWriter(entry.Open()))
                {
                    writer.Write(source);
                }
            }
        }

        Console.WriteLine($"Exported {loadedDoc.VbaProject.Modules.Count} VBA module(s) to \"{zipPath}\".");
    }
}
