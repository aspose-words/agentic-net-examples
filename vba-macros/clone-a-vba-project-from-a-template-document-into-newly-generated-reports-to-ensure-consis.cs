using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file paths for the template and the generated report.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string templatePath = Path.Combine(artifactsDir, "Template.docm");
        string reportPath = Path.Combine(artifactsDir, "Report.docm");

        // -----------------------------------------------------------------
        // 1. Create a template document that contains a VBA project.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();

        // Create a new VBA project and assign a name.
        VbaProject templateProject = new VbaProject
        {
            Name = "TemplateProject"
        };
        templateDoc.VbaProject = templateProject;

        // Create a procedural module with a simple macro.
        VbaModule macroModule = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from the template macro!""
End Sub"
        };
        // Add the module to the VBA project.
        templateDoc.VbaProject.Modules.Add(macroModule);

        // Save the template as a macro‑enabled document.
        templateDoc.Save(templatePath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // 2. Load the template (optional – we already have it in memory).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Ensure the template actually contains macros before cloning.
        if (!loadedTemplate.HasMacros || loadedTemplate.VbaProject == null)
        {
            throw new InvalidOperationException("The template does not contain a VBA project.");
        }

        // -----------------------------------------------------------------
        // 3. Create a new report document and clone the VBA project from the template.
        // -----------------------------------------------------------------
        Document reportDoc = new Document();

        // Clone the VBA project from the template.
        VbaProject clonedProject = loadedTemplate.VbaProject.Clone();

        // Assign the cloned project to the report document.
        reportDoc.VbaProject = clonedProject;

        // Save the report as a macro‑enabled document.
        reportDoc.Save(reportPath, SaveFormat.Docm);
    }
}
