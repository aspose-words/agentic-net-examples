using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // ---------- Create an initial VBA project ----------
        VbaProject originalProject = new VbaProject();
        originalProject.Name = "OriginalProject";

        VbaModule originalModule = new VbaModule();
        originalModule.Name = "OriginalModule";
        originalModule.Type = VbaModuleType.ProceduralModule;
        originalModule.SourceCode = "Sub Hello()\n    MsgBox \"Hello from original\"\nEnd Sub";

        originalProject.Modules.Add(originalModule);

        // Assign the original VBA project to the document.
        doc.VbaProject = originalProject;

        // Save the document with the original VBA project.
        doc.Save("Original.docm");

        // ---------- Create a pre‑configured VBA project template ----------
        VbaProject templateProject = new VbaProject();
        templateProject.Name = "TemplateProject";

        VbaModule templateModule = new VbaModule();
        templateModule.Name = "TemplateModule";
        templateModule.Type = VbaModuleType.ProceduralModule;
        templateModule.SourceCode = "Sub Hello()\n    MsgBox \"Hello from template\"\nEnd Sub";

        templateProject.Modules.Add(templateModule);

        // Replace the document's VBA project with the template project.
        doc.VbaProject = templateProject;

        // Save the document after replacement.
        doc.Save("Replaced.docm");
    }
}
