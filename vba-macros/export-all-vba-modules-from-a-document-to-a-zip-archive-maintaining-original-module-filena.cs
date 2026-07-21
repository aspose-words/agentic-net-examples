using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Vba;

public class ExportVbaModules
{
    public static void Main()
    {
        // Define file and folder paths.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        string docPath = Path.Combine(workDir, "SampleDocument.docm");
        string zipPath = Path.Combine(workDir, "VbaModules.zip");
        string tempModulesDir = Path.Combine(workDir, "ModulesTemp");

        // -----------------------------------------------------------------
        // 1. Create a macro‑enabled document with a few VBA modules.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        project.Name = "SampleProject";
        doc.VbaProject = project;

        // Add sample modules.
        for (int i = 1; i <= 3; i++)
        {
            VbaModule module = new VbaModule();
            module.Name = $"Module{i}";
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = $"Sub Macro{i}()\n    MsgBox \"Hello from Module{i}\"\nEnd Sub";
            project.Modules.Add(module);
        }

        // Save the document in macro‑enabled format.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Load the document and export each VBA module to a file.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
        {
            Console.WriteLine("The document does not contain any VBA macros.");
            return;
        }

        // Prepare temporary folder for module files.
        if (Directory.Exists(tempModulesDir))
            Directory.Delete(tempModulesDir, true);
        Directory.CreateDirectory(tempModulesDir);

        // Export each module preserving its original name.
        foreach (VbaModule module in loadedDoc.VbaProject.Modules)
        {
            string moduleName = module.Name ?? "UnnamedModule";
            string source = module.SourceCode ?? string.Empty;
            string fileName = $"{moduleName}.bas"; // .bas is a common VBA module extension
            string filePath = Path.Combine(tempModulesDir, fileName);
            File.WriteAllText(filePath, source);
        }

        // -----------------------------------------------------------------
        // 3. Create a ZIP archive containing the exported modules.
        // -----------------------------------------------------------------
        if (File.Exists(zipPath))
            File.Delete(zipPath);

        ZipFile.CreateFromDirectory(tempModulesDir, zipPath);

        // Cleanup temporary files.
        Directory.Delete(tempModulesDir, true);

        // Indicate completion.
        Console.WriteLine($"Exported {loadedDoc.VbaProject.Modules.Count} VBA module(s) to \"{zipPath}\".");
    }
}
