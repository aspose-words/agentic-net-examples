using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class ReplaceVbaProjectExample
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docm");
        string replacedPath = Path.Combine(Directory.GetCurrentDirectory(), "Replaced.docm");

        // -------------------------------------------------
        // Step 1: Create a blank document with an initial VBA project.
        // -------------------------------------------------
        Document originalDoc = new Document();

        // Create the initial VBA project.
        VbaProject initialProject = new VbaProject
        {
            Name = "InitialProject"
        };

        // Add a simple module to the initial project.
        VbaModule initialModule = new VbaModule
        {
            Name = "InitialModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from the initial project!\"\nEnd Sub"
        };
        initialProject.Modules.Add(initialModule);

        // Assign the VBA project to the document.
        originalDoc.VbaProject = initialProject;

        // Save the document as a macro‑enabled file.
        originalDoc.Save(originalPath, SaveFormat.Docm);

        // -------------------------------------------------
        // Step 2: Prepare a pre‑configured VBA project template.
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
            SourceCode = "Sub TemplateMacro1()\n    MsgBox \"This is template macro 1.\"\nEnd Sub"
        };
        templateProject.Modules.Add(templateModule1);

        // Second module in the template.
        VbaModule templateModule2 = new VbaModule
        {
            Name = "TemplateModule2",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub TemplateMacro2()\n    MsgBox \"This is template macro 2.\"\nEnd Sub"
        };
        templateProject.Modules.Add(templateModule2);

        // -------------------------------------------------
        // Step 3: Load the original document and replace its VBA project.
        // -------------------------------------------------
        Document docToReplace = new Document(originalPath);

        // Replace the existing VBA project with the template project.
        docToReplace.VbaProject = templateProject;

        // Save the modified document.
        docToReplace.Save(replacedPath, SaveFormat.Docm);

        // -------------------------------------------------
        // Step 4: Simple validation – output module names to console.
        // -------------------------------------------------
        Document validationDoc = new Document(replacedPath);
        Console.WriteLine($"Document '{replacedPath}' now contains VBA project '{validationDoc.VbaProject.Name}'.");
        foreach (VbaModule module in validationDoc.VbaProject.Modules)
        {
            Console.WriteLine($"Module: {module.Name}");
        }
    }
}
