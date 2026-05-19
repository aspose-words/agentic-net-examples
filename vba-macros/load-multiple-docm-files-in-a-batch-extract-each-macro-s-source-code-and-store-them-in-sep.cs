using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define folders for sample documents and extracted macros.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string macrosDir = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedMacros");

        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(macrosDir);

        // Create sample macro‑enabled documents if they do not already exist.
        CreateSampleDocument(Path.Combine(dataDir, "Sample1.docm"), "ModuleA", "Sub HelloWorld()\n    MsgBox \"Hello from Sample1!\"\nEnd Sub");
        CreateSampleDocument(Path.Combine(dataDir, "Sample2.docm"), "ModuleB", "Function AddNumbers(a As Integer, b As Integer) As Integer\n    AddNumbers = a + b\nEnd Function");

        // Process each .docm file in the data directory.
        foreach (string docPath in Directory.GetFiles(dataDir, "*.docm"))
        {
            // Load the document.
            Document doc = new Document(docPath);

            // Ensure the document contains a VBA project.
            if (!doc.HasMacros || doc.VbaProject == null)
                continue;

            // Iterate through all VBA modules.
            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                // Guard against null source code.
                string source = module.SourceCode ?? string.Empty;

                // Build a filename that identifies the source document and module.
                string macroFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_{module.Name}.bas";
                string macroFilePath = Path.Combine(macrosDir, macroFileName);

                // Write the macro source to a file.
                File.WriteAllText(macroFilePath, source);
            }
        }
    }

    // Helper method to create a macro‑enabled document with a single module.
    private static void CreateSampleDocument(string filePath, string moduleName, string moduleSource)
    {
        // If the file already exists we skip creation to avoid overwriting.
        if (File.Exists(filePath))
            return;

        // Create a blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = project;

        // Create a new module, set its properties, and add it to the project.
        VbaModule module = new VbaModule
        {
            Name = moduleName,
            Type = VbaModuleType.ProceduralModule,
            SourceCode = moduleSource
        };
        doc.VbaProject.Modules.Add(module);

        // Save the document in macro‑enabled format.
        doc.Save(filePath);
    }
}
