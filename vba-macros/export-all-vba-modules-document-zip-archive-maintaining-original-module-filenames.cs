using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Vba;

class ExportVbaModulesToZip
{
    static void Main()
    {
        // Path to the source document that contains VBA macros.
        string docsPath = @"C:\Docs\Macro.docm";

        // Path where the resulting ZIP archive will be saved.
        string zipPath = @"C:\Docs\VbaModules.zip";

        // Load the document (uses the provided Document load rule).
        Document doc = new Document(docsPath);

        // Ensure the document actually contains a VBA project.
        if (doc.VbaProject == null || doc.VbaProject.Modules == null)
        {
            Console.WriteLine("The document does not contain any VBA modules.");
            return;
        }

        // Create the ZIP archive (ZIP creation is not covered by a specific rule, so we use standard .NET APIs).
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create))
        {
            // Iterate through each VBA module in the project.
            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                // Preserve the original module name as the file name.
                // Append a .bas extension for procedural modules; other types could use different extensions if needed.
                string entryName = $"{module.Name}.bas";

                // Create a new entry in the ZIP archive.
                ZipArchiveEntry entry = archive.CreateEntry(entryName);

                // Write the module's source code into the entry.
                using (StreamWriter writer = new StreamWriter(entry.Open()))
                {
                    writer.Write(module.SourceCode);
                }
            }
        }

        Console.WriteLine($"Exported {doc.VbaProject.Modules.Count} VBA module(s) to \"{zipPath}\".");
    }
}
