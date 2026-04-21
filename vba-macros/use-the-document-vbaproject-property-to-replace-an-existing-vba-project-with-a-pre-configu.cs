using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docm");
        string replacedPath = Path.Combine(Directory.GetCurrentDirectory(), "Replaced.docm");

        // -------------------------------------------------
        // Step 1: Create a sample document with an initial VBA project.
        // -------------------------------------------------
        Document originalDoc = new Document();

        // Create a VBA project and add a single module.
        VbaProject originalProject = new VbaProject
        {
            Name = "OriginalProject"
        };

        VbaModule originalModule = new VbaModule
        {
            Name = "OriginalModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloOriginal()\n    MsgBox \"Hello from original macro!\"\nEnd Sub"
        };

        originalProject.Modules.Add(originalModule);
        originalDoc.VbaProject = originalProject;

        // Save the document in a macro‑enabled format.
        originalDoc.Save(originalPath, SaveFormat.Docm);

        // -------------------------------------------------
        // Step 2: Load the document (simulating an existing file).
        // -------------------------------------------------
        Document doc = new Document(originalPath);

        // -------------------------------------------------
        // Step 3: Create a pre‑configured VBA project template.
        // -------------------------------------------------
        VbaProject templateProject = new VbaProject
        {
            Name = "TemplateProject"
        };

        // First module in the template.
        VbaModule templateModule1 = new VbaModule
        {
            Name = "TemplateModule1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloTemplate1()\n    MsgBox \"Hello from template macro 1!\"\nEnd Sub"
        };
        templateProject.Modules.Add(templateModule1);

        // Second module in the template.
        VbaModule templateModule2 = new VbaModule
        {
            Name = "TemplateModule2",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloTemplate2()\n    MsgBox \"Hello from template macro 2!\"\nEnd Sub"
        };
        templateProject.Modules.Add(templateModule2);

        // -------------------------------------------------
        // Step 4: Replace the document's VBA project with the template.
        // -------------------------------------------------
        doc.VbaProject = templateProject;

        // -------------------------------------------------
        // Step 5: Save the modified document.
        // -------------------------------------------------
        doc.Save(replacedPath, SaveFormat.Docm);

        // -------------------------------------------------
        // Step 6: Simple validation output.
        // -------------------------------------------------
        Console.WriteLine($"Original document saved to: {originalPath}");
        Console.WriteLine($"Replaced document saved to: {replacedPath}");

        // Load the replaced document to verify the replacement.
        Document verificationDoc = new Document(replacedPath);
        Console.WriteLine($"Has macros: {verificationDoc.HasMacros}");
        Console.WriteLine($"VBA project name: {verificationDoc.VbaProject?.Name}");
        Console.WriteLine($"Number of modules: {verificationDoc.VbaProject?.Modules?.Count}");

        // List module names and a snippet of their source code.
        if (verificationDoc.VbaProject != null)
        {
            foreach (VbaModule module in verificationDoc.VbaProject.Modules)
            {
                string codeSnippet = module.SourceCode?.Split('\n')[0] ?? string.Empty;
                Console.WriteLine($"Module: {module.Name}, First line of code: {codeSnippet}");
            }
        }
    }
}
