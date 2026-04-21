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

        // Path for the sample macro‑enabled document.
        string docPath = Path.Combine(outputDir, "Sample.docm");

        // Create a blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        VbaProject vbaProject = new VbaProject { Name = "SampleProject" };
        doc.VbaProject = vbaProject;

        // Add a few VBA modules with simple macros.
        for (int i = 1; i <= 3; i++)
        {
            VbaModule module = new VbaModule
            {
                Name = $"Module{i}",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = $"Sub Macro{i}()\n    MsgBox \"Hello from Macro{i}\"\nEnd Sub"
            };
            doc.VbaProject.Modules.Add(module);
        }

        // Save the document in a macro‑enabled format.
        doc.Save(docPath);

        // Export all VBA modules to a ZIP archive, preserving original module names.
        string zipPath = Path.Combine(outputDir, "VbaModules.zip");
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create))
        {
            // Reload the document to ensure modules are persisted (optional).
            Document loadedDoc = new Document(docPath);

            // Guard against missing VBA project.
            if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
            {
                foreach (VbaModule module in loadedDoc.VbaProject.Modules)
                {
                    // Use the module name as the file name inside the ZIP.
                    string entryName = $"{module.Name}.bas";

                    // Ensure source code is not null.
                    string source = module.SourceCode ?? string.Empty;

                    // Create a new entry and write the source code.
                    ZipArchiveEntry entry = archive.CreateEntry(entryName);
                    using (Stream entryStream = entry.Open())
                    using (StreamWriter writer = new StreamWriter(entryStream))
                    {
                        writer.Write(source);
                    }
                }
            }
        }

        // The program finishes automatically; no user interaction required.
    }
}
