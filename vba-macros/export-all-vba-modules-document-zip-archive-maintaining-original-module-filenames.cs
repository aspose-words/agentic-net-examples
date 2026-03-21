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

        // Verify that the source document exists.
        if (!File.Exists(docsPath))
        {
            Console.WriteLine($"Source file not found: {docsPath}");
            return;
        }

        // Ensure the output directory exists.
        string zipDirectory = Path.GetDirectoryName(zipPath);
        if (!Directory.Exists(zipDirectory))
        {
            Directory.CreateDirectory(zipDirectory);
        }

        // Load the document.
        Document doc = new Document(docsPath);

        // Ensure the document actually contains a VBA project.
        if (doc.VbaProject == null || doc.VbaProject.Modules == null || doc.VbaProject.Modules.Count == 0)
        {
            Console.WriteLine("The document does not contain any VBA modules.");
            return;
        }

        // Create the ZIP archive.
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create))
        {
            // Iterate through each VBA module in the project.
            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                // Preserve the original module name as the file name.
                // Append a .bas extension for procedural modules.
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
