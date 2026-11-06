using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main(string[] args)
    {
        // Folder for sample macro-enabled documents
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsFolder);

        // Create sample DOCM files with VBA modules
        CreateSampleDocM(Path.Combine(docsFolder, "Sample1.docm"), "ModuleA",
            "Sub Greet()\n    MsgBox \"Hello from Sample1\"\nEnd Sub");
        CreateSampleDocM(Path.Combine(docsFolder, "Sample2.docm"), "ModuleB",
            "Sub ShowInfo()\n    MsgBox \"Info from Sample2\"\nEnd Sub");

        // Folder for extracted macro source files
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedMacros");
        Directory.CreateDirectory(outputFolder);

        // Process each DOCM file in the Docs folder
        foreach (string docPath in Directory.GetFiles(docsFolder, "*.docm"))
        {
            // Load the macro-enabled document
            Document doc = new Document(docPath);

            // Ensure the document contains a VBA project
            VbaProject vbaProject = doc.VbaProject;
            if (vbaProject == null)
                continue; // No macros to extract

            // Iterate through all VBA modules
            foreach (VbaModule module in vbaProject.Modules)
            {
                // Guard against null source code
                string source = module.SourceCode ?? string.Empty;

                // Build a file name: <DocumentName>_<ModuleName>.bas
                string docFileName = Path.GetFileNameWithoutExtension(docPath);
                string macroFileName = $"{docFileName}_{module.Name}.bas";
                string macroFilePath = Path.Combine(outputFolder, macroFileName);

                // Write the source code to the file
                File.WriteAllText(macroFilePath, source);
            }
        }
    }

    // Helper method to create a macro-enabled DOCM with a single VBA module
    private static void CreateSampleDocM(string filePath, string moduleName, string moduleSource)
    {
        // Create a blank document
        Document doc = new Document();

        // Initialize a new VBA project
        VbaProject vbaProject = new VbaProject();

        // Create a VBA module with the provided name and source code
        VbaModule vbaModule = new VbaModule();
        vbaModule.Name = moduleName;
        vbaModule.SourceCode = moduleSource;

        // Add the module to the project
        vbaProject.Modules.Add(vbaModule);

        // Assign the VBA project to the document
        doc.VbaProject = vbaProject;

        // Save as a macro-enabled document
        doc.Save(filePath, SaveFormat.Docm);
    }
}
