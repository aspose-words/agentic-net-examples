using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Directory for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Create a sample macro‑enabled document with an initial VBA project.
        // -----------------------------------------------------------------
        Document originalDoc = new Document();

        // Create the original VBA project.
        VbaProject originalProject = new VbaProject();
        originalProject.Name = "OriginalProject";

        // Add a module to the original project.
        VbaModule originalModule = new VbaModule();
        originalModule.Name = "OriginalModule";
        originalModule.Type = VbaModuleType.ProceduralModule;
        originalModule.SourceCode = "Sub Hello()\n    MsgBox \"Hello from original\"\nEnd Sub";

        originalProject.Modules.Add(originalModule);
        originalDoc.VbaProject = originalProject;

        // Save the document as a macro‑enabled file.
        string originalPath = Path.Combine(outputDir, "Original.docm");
        originalDoc.Save(originalPath);

        // -----------------------------------------------------------------
        // Step 2: Prepare a pre‑configured VBA project template.
        // -----------------------------------------------------------------
        VbaProject templateProject = new VbaProject();
        templateProject.Name = "TemplateProject";

        VbaModule templateModule = new VbaModule();
        templateModule.Name = "TemplateModule";
        templateModule.Type = VbaModuleType.ProceduralModule;
        templateModule.SourceCode = "Sub Hello()\n    MsgBox \"Hello from template\"\nEnd Sub";

        templateProject.Modules.Add(templateModule);

        // -----------------------------------------------------------------
        // Step 3: Load the original document and replace its VBA project.
        // -----------------------------------------------------------------
        Document docToReplace = new Document(originalPath);

        // Replace the existing VBA project with the template project.
        docToReplace.VbaProject = templateProject;

        // Save the result.
        string replacedPath = Path.Combine(outputDir, "Replaced.docm");
        docToReplace.Save(replacedPath);

        // -----------------------------------------------------------------
        // Step 4: Simple validation – output module information to the console.
        // -----------------------------------------------------------------
        Console.WriteLine("After replacement, VBA project name: " + docToReplace.VbaProject.Name);
        Console.WriteLine("Number of modules: " + docToReplace.VbaProject.Modules.Count);
        foreach (VbaModule module in docToReplace.VbaProject.Modules)
        {
            Console.WriteLine($"Module: {module.Name}");
            Console.WriteLine("Source code:");
            Console.WriteLine(module.SourceCode);
        }
    }
}
