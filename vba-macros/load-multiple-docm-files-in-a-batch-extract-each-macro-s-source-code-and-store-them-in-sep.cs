using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Directory to store sample documents and extracted macro files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create sample macro-enabled documents.
        CreateSampleDocM(Path.Combine(artifactsDir, "Sample1.docm"), "MacroOne", "Sub MacroOne()\n    MsgBox \"Hello from MacroOne\"\nEnd Sub");
        CreateSampleDocM(Path.Combine(artifactsDir, "Sample2.docm"), "MacroTwo", "Sub MacroTwo()\n    MsgBox \"Hello from MacroTwo\"\nEnd Sub");

        // Process each .docm file in the artifacts directory.
        string[] docmFiles = Directory.GetFiles(artifactsDir, "*.docm");
        foreach (string docmPath in docmFiles)
        {
            // Load the document.
            Document doc = new Document(docmPath);

            // Skip files without macros.
            if (!doc.HasMacros || doc.VbaProject == null)
                continue;

            VbaProject vbaProject = doc.VbaProject;
            VbaModuleCollection modules = vbaProject.Modules;

            // Extract each module's source code.
            foreach (VbaModule module in modules)
            {
                // Guard against null source code.
                string source = module.SourceCode ?? string.Empty;

                // Build a filename for the extracted macro.
                string macroFileName = $"{Path.GetFileNameWithoutExtension(docmPath)}_{module.Name}.bas";
                string macroFilePath = Path.Combine(artifactsDir, macroFileName);

                // Write the source code to a file.
                File.WriteAllText(macroFilePath, source);
            }
        }
    }

    // Helper method to create a macro-enabled document with a single module.
    private static void CreateSampleDocM(string filePath, string moduleName, string sourceCode)
    {
        // Create a blank document.
        Document doc = new Document();

        // Ensure a VBA project exists.
        VbaProject project = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = project;

        // Create a new module and set its properties.
        VbaModule module = new VbaModule
        {
            Name = moduleName,
            Type = VbaModuleType.ProceduralModule,
            SourceCode = sourceCode
        };

        // Add the module to the project.
        doc.VbaProject.Modules.Add(module);

        // Save as a macro-enabled document.
        doc.Save(filePath);
    }
}
