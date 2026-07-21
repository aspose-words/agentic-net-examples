using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Directory to store sample documents and extracted macro files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create sample macro-enabled documents.
        CreateSampleDocument(Path.Combine(workDir, "Sample1.docm"), "SampleProject1",
            ("ModuleA", "Sub MacroA()\n    MsgBox \"Hello from MacroA\"\nEnd Sub"));
        CreateSampleDocument(Path.Combine(workDir, "Sample2.docm"), "SampleProject2",
            ("ModuleB", "Sub MacroB()\n    MsgBox \"Hello from MacroB\"\nEnd Sub"),
            ("ModuleC", "Sub MacroC()\n    MsgBox \"Hello from MacroC\"\nEnd Sub"));

        // Directory to store extracted macro source files.
        string macrosDir = Path.Combine(workDir, "ExtractedMacros");
        Directory.CreateDirectory(macrosDir);

        // Load each .docm file, extract macro source code, and save to separate files.
        foreach (string docPath in Directory.GetFiles(workDir, "*.docm"))
        {
            Document doc = new Document(docPath);

            // Ensure the document contains a VBA project.
            if (!doc.HasMacros || doc.VbaProject == null)
                continue;

            VbaProject vbaProject = doc.VbaProject;
            foreach (VbaModule module in vbaProject.Modules)
            {
                // Guard against null source code.
                string source = module.SourceCode ?? string.Empty;

                // Build a filename that includes the original document name and module name.
                string macroFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_{module.Name}.bas";
                string macroFilePath = Path.Combine(macrosDir, macroFileName);

                File.WriteAllText(macroFilePath, source);
            }
        }
    }

    // Helper method to create a macro-enabled document with one or more modules.
    private static void CreateSampleDocument(string filePath, string projectName, params (string moduleName, string sourceCode)[] modules)
    {
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject
        {
            Name = projectName
        };
        doc.VbaProject = project;

        // Add each provided module to the VBA project.
        foreach (var (moduleName, sourceCode) in modules)
        {
            VbaModule vbaModule = new VbaModule
            {
                Name = moduleName,
                Type = VbaModuleType.ProceduralModule,
                SourceCode = sourceCode
            };
            doc.VbaProject.Modules.Add(vbaModule);
        }

        // Save as a macro-enabled document.
        doc.Save(filePath);
    }
}
