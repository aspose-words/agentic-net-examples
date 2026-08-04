using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class ExtractVbaMacros
{
    public static void Main()
    {
        // Base directory for generated files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(baseDir);

        // Create sample macro‑enabled documents.
        CreateSampleDocument(Path.Combine(baseDir, "Sample1.docm"), "SampleProject1",
            ("ModuleA", "Sub Hello()\n    MsgBox \"Hello from ModuleA\"\nEnd Sub"));
        CreateSampleDocument(Path.Combine(baseDir, "Sample2.docm"), "SampleProject2",
            ("ModuleB", "Function Add(a As Integer, b As Integer) As Integer\n    Add = a + b\nEnd Function"));

        // Directory where extracted macro source files will be saved.
        string macrosDir = Path.Combine(baseDir, "ExtractedMacros");
        Directory.CreateDirectory(macrosDir);

        // Process each .docm file in the base directory.
        foreach (string docPath in Directory.GetFiles(baseDir, "*.docm"))
        {
            // Load the document.
            Document doc = new Document(docPath);

            // Ensure the document actually contains a VBA project.
            if (!doc.HasMacros || doc.VbaProject == null)
                continue;

            VbaProject vbaProject = doc.VbaProject;
            VbaModuleCollection modules = vbaProject.Modules;

            // Extract each module's source code.
            foreach (VbaModule module in modules)
            {
                // Guard against null source code.
                string source = module.SourceCode ?? string.Empty;

                // Build a filename that identifies the source document and module.
                string docFileName = Path.GetFileNameWithoutExtension(docPath);
                string macroFileName = $"{docFileName}_{module.Name}.bas";
                string macroFilePath = Path.Combine(macrosDir, macroFileName);

                // Write the source code to a file.
                File.WriteAllText(macroFilePath, source);
            }
        }
    }

    // Helper method to create a macro‑enabled document with a single module.
    private static void CreateSampleDocument(string filePath, string projectName, (string Name, string Code) moduleInfo)
    {
        // Create a blank document.
        Document doc = new Document();

        // Create a new VBA project and assign a name.
        VbaProject project = new VbaProject();
        project.Name = projectName;
        doc.VbaProject = project;

        // Create a new module, set its properties, and add it to the project.
        VbaModule module = new VbaModule();
        module.Name = moduleInfo.Name;
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = moduleInfo.Code;
        doc.VbaProject.Modules.Add(module);

        // Save as a macro‑enabled document.
        doc.Save(filePath);
    }
}
