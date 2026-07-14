using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class ExtractVbaMacros
{
    public static void Main()
    {
        // Base folder for sample documents and extracted macro files.
        string baseFolder = Path.Combine(Directory.GetCurrentDirectory(), "MacroDemo");
        string docsFolder = Path.Combine(baseFolder, "Docs");
        string outputFolder = Path.Combine(baseFolder, "ExtractedMacros");

        Directory.CreateDirectory(docsFolder);
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Step 1: Create sample macro‑enabled documents (DOCM) if they do not exist.
        // -----------------------------------------------------------------
        CreateSampleDocument(Path.Combine(docsFolder, "Sample1.docm"), "ModuleA", "Sub HelloWorld()\n    MsgBox \"Hello from Sample1!\"\nEnd Sub");
        CreateSampleDocument(Path.Combine(docsFolder, "Sample2.docm"), "ModuleB", "Function AddNumbers(a As Integer, b As Integer) As Integer\n    AddNumbers = a + b\nEnd Function");

        // -----------------------------------------------------------------
        // Step 2: Load each DOCM file, extract VBA modules, and save their source code.
        // -----------------------------------------------------------------
        foreach (string docPath in Directory.GetFiles(docsFolder, "*.docm"))
        {
            // Load the document.
            Document doc = new Document(docPath);

            // Verify that the document actually contains a VBA project.
            if (!doc.HasMacros || doc.VbaProject == null)
                continue; // No macros to extract.

            VbaProject vbaProject = doc.VbaProject;
            VbaModuleCollection modules = vbaProject.Modules;

            // Iterate through all modules in the project.
            foreach (VbaModule module in modules)
            {
                // Guard against null source code.
                string source = module.SourceCode ?? string.Empty;

                // Build a unique file name for the extracted macro.
                string macroFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_{module.Name}.bas";
                string macroFilePath = Path.Combine(outputFolder, macroFileName);

                // Write the source code to the file.
                File.WriteAllText(macroFilePath, source);
            }
        }
    }

    // Helper method to create a macro‑enabled document with a single VBA module.
    private static void CreateSampleDocument(string filePath, string moduleName, string moduleSource)
    {
        // If the file already exists we skip creation to avoid overwriting.
        if (File.Exists(filePath))
            return;

        // Create a blank document.
        Document doc = new Document();

        // Create a new VBA project and assign a name.
        VbaProject project = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = project;

        // Create a new VBA module, set its properties, and add it to the project.
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
