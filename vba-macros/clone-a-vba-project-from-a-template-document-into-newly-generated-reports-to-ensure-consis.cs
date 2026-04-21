using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class CloneVbaProjectExample
{
    public static void Main()
    {
        // Define file names in the current directory.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docm");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docm");

        // -------------------------------------------------
        // 1. Create a template document with a VBA project.
        // -------------------------------------------------
        Document templateDoc = new Document();

        // Create a new VBA project and give it a name.
        VbaProject templateProject = new VbaProject
        {
            Name = "TemplateProject"
        };

        // Create a procedural module that contains a simple macro.
        VbaModule macroModule = new VbaModule
        {
            Name = "HelloModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"Sub HelloWorld()
    MsgBox ""Hello from the template!""
End Sub"
        };

        // Add the module to the VBA project.
        templateProject.Modules.Add(macroModule);

        // Attach the VBA project to the template document.
        templateDoc.VbaProject = templateProject;

        // Save the template as a macro‑enabled document.
        templateDoc.Save(templatePath, SaveFormat.Docm);

        // -------------------------------------------------
        // 2. Load the template and clone its VBA project.
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Ensure the template actually contains a VBA project.
        if (!loadedTemplate.HasMacros || loadedTemplate.VbaProject == null)
        {
            throw new InvalidOperationException("The template does not contain a VBA project.");
        }

        // Clone the VBA project from the template.
        VbaProject clonedProject = loadedTemplate.VbaProject.Clone();

        // -------------------------------------------------
        // 3. Create a new report document and assign the cloned project.
        // -------------------------------------------------
        Document reportDoc = new Document();

        // Assign the cloned VBA project to the report.
        reportDoc.VbaProject = clonedProject;

        // Save the report as a macro‑enabled document.
        reportDoc.Save(reportPath, SaveFormat.Docm);

        // -------------------------------------------------
        // 4. Simple validation (optional).
        // -------------------------------------------------
        // Verify that the report now has the macro.
        if (reportDoc.HasMacros && reportDoc.VbaProject != null)
        {
            VbaModule reportModule = reportDoc.VbaProject.Modules["HelloModule"];
            if (reportModule != null && !string.IsNullOrEmpty(reportModule.SourceCode))
            {
                // Output a confirmation to the console.
                Console.WriteLine("Macro successfully cloned to the report document.");
            }
        }
    }
}
