using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a source document that already contains a VBA project.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();

        VbaProject sourceProject = new VbaProject();
        sourceProject.Name = "SourceProject";

        VbaModule sourceModule = new VbaModule();
        sourceModule.Name = "SourceModule";
        sourceModule.Type = VbaModuleType.ProceduralModule;
        sourceModule.SourceCode = "Sub Hello()\n    MsgBox \"Hello from source\"\nEnd Sub";

        sourceProject.Modules.Add(sourceModule);
        sourceDoc.VbaProject = sourceProject;

        const string sourcePath = "Source.docm";
        sourceDoc.Save(sourcePath); // Save as macro‑enabled document.

        // ---------------------------------------------------------------
        // 2. Prepare a pre‑configured VBA project template to replace with.
        // ---------------------------------------------------------------
        VbaProject templateProject = new VbaProject();
        templateProject.Name = "TemplateProject";

        VbaModule templateModule = new VbaModule();
        templateModule.Name = "TemplateModule";
        templateModule.Type = VbaModuleType.ProceduralModule;
        templateModule.SourceCode = "Sub Hello()\n    MsgBox \"Hello from template\"\nEnd Sub";

        templateProject.Modules.Add(templateModule);

        // ---------------------------------------------------------------
        // 3. Load the existing document and replace its VBA project.
        // ---------------------------------------------------------------
        Document doc = new Document(sourcePath);

        // Replace the VBA project with a clone of the template project.
        doc.VbaProject = templateProject.Clone();

        // ---------------------------------------------------------------
        // 4. Save the modified document.
        // ---------------------------------------------------------------
        const string resultPath = "Result.docm";
        doc.Save(resultPath); // Must be saved as .docm to retain macros.

        // ---------------------------------------------------------------
        // 5. Simple validation: ensure the replacement succeeded.
        // ---------------------------------------------------------------
        VbaModule replacedModule = doc.VbaProject?.Modules["TemplateModule"];
        Console.WriteLine(replacedModule != null
            ? "VBA project replacement succeeded."
            : "VBA project replacement failed.");

        Console.WriteLine("Replaced module source code:");
        Console.WriteLine(replacedModule?.SourceCode ?? string.Empty);
    }
}
