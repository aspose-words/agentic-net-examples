using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Saving;
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

        // Access the module collection.
        VbaModuleCollection modules = doc.VbaProject.Modules;

        // Create ModuleOne.
        VbaModule moduleOne = new VbaModule();
        moduleOne.Name = "ModuleOne";
        moduleOne.SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from ModuleOne\"\nEnd Sub";
        modules.Add(moduleOne);

        // Create ModuleTwo.
        VbaModule moduleTwo = new VbaModule();
        moduleTwo.Name = "ModuleTwo";
        moduleTwo.SourceCode = "Sub ShowDate()\n    MsgBox \"Today is \" & Date\nEnd Sub";
        modules.Add(moduleTwo);

        // Save the document as a macro‑enabled file.
        const string docPath = "Sample.docm";
        doc.Save(docPath, SaveFormat.Docm);

        // Reload the document to simulate a fresh load.
        Document loadedDoc = new Document(docPath);

        // Prepare the ZIP archive to store VBA modules.
        const string zipPath = "VbaModules.zip";
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
        {
            // Iterate through all VBA modules in the project.
            foreach (VbaModule module in loadedDoc.VbaProject.Modules)
            {
                // Use the module name as the file name, adding .bas extension.
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

        // Program ends.
    }
}
