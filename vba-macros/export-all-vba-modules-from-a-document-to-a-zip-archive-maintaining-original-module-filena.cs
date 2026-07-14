using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the temporary macro-enabled document.
        string docPath = Path.Combine(outputDir, "Sample.docm");

        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = project;

        // Add a couple of VBA modules with sample code.
        VbaModule module1 = new VbaModule
        {
            Name = "ModuleOne",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from ModuleOne\"\nEnd Sub"
        };
        VbaModule module2 = new VbaModule
        {
            Name = "ModuleTwo",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub GoodbyeWorld()\n    MsgBox \"Goodbye from ModuleTwo\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(module1);
        doc.VbaProject.Modules.Add(module2);

        // Save the document in macro-enabled format.
        doc.Save(docPath);

        // Reload the document to ensure we work with persisted data.
        Document loadedDoc = new Document(docPath);

        // Verify that the document contains a VBA project.
        if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
        {
            Console.WriteLine("No VBA project found in the document.");
            return;
        }

        // Prepare the ZIP archive path.
        string zipPath = Path.Combine(outputDir, "VbaModules.zip");

        // Create the ZIP archive and add each VBA module as a separate entry.
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create))
        {
            foreach (VbaModule vbaModule in loadedDoc.VbaProject.Modules)
            {
                // Use the module name with .bas extension to preserve original filenames.
                string entryName = $"{vbaModule.Name}.bas";

                // Ensure source code is not null.
                string source = vbaModule.SourceCode ?? string.Empty;

                ZipArchiveEntry entry = archive.CreateEntry(entryName);
                using (StreamWriter writer = new StreamWriter(entry.Open()))
                {
                    writer.Write(source);
                }
            }
        }

        Console.WriteLine($"Exported {loadedDoc.VbaProject.Modules.Count} VBA module(s) to \"{zipPath}\".");
    }
}
