using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(outputDir, "Template.docm");
        string reportPath = Path.Combine(outputDir, "Report.docm");

        // -------------------------------------------------
        // 1. Create a template document that contains a VBA project.
        // -------------------------------------------------
        Document templateDoc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject templateProject = new VbaProject
        {
            Name = "TemplateProject"
        };
        templateDoc.VbaProject = templateProject;

        // Create a VBA module with a simple macro.
        VbaModule macroModule = new VbaModule
        {
            Name = "StandardModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from template!""
End Sub
"
        };

        // Add the module to the project's module collection.
        templateDoc.VbaProject.Modules.Add(macroModule);

        // Save the template as a macro‑enabled document.
        templateDoc.Save(templatePath, SaveFormat.Docm);

        // -------------------------------------------------
        // 2. Load the template and clone its VBA project.
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        VbaProject clonedProject = loadedTemplate.VbaProject.Clone();

        // -------------------------------------------------
        // 3. Create a new report document and attach the cloned VBA project.
        // -------------------------------------------------
        Document reportDoc = new Document();
        reportDoc.VbaProject = clonedProject;

        // Optional: modify the cloned macro (e.g., add a comment) safely.
        VbaModule existingModule = reportDoc.VbaProject.Modules["StandardModule"];
        if (existingModule != null)
        {
            // Guard against null source code.
            string source = existingModule.SourceCode ?? string.Empty;
            existingModule.SourceCode = source + "\r\n' Cloned from template";
        }

        // Save the report as a macro‑enabled document.
        reportDoc.Save(reportPath, SaveFormat.Docm);
    }
}
