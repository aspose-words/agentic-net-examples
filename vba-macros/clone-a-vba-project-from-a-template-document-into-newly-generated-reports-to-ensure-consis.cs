using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // ---------- Create a template document with a VBA project ----------
        Document templateDoc = new Document();

        // Create a new VBA project and give it a name.
        VbaProject templateProject = new VbaProject();
        templateProject.Name = "TemplateProject";
        templateDoc.VbaProject = templateProject;

        // Create a VBA module containing a simple macro.
        VbaModule module = new VbaModule();
        module.Name = "SampleModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from template!""
End Sub
";

        // Add the module to the project's module collection.
        templateDoc.VbaProject.Modules.Add(module);

        // Save the template as a macro‑enabled document.
        string templatePath = Path.Combine(outputDir, "Template.docm");
        templateDoc.Save(templatePath, SaveFormat.Docm);

        // ---------- Load the template and clone its VBA project ----------
        Document loadedTemplate = new Document(templatePath);
        VbaProject clonedProject = loadedTemplate.VbaProject.Clone();

        // ---------- Create a new report document and attach the cloned VBA project ----------
        Document reportDoc = new Document();
        reportDoc.VbaProject = clonedProject;

        // Save the report as a macro‑enabled document.
        string reportPath = Path.Combine(outputDir, "Report.docm");
        reportDoc.Save(reportPath, SaveFormat.Docm);
    }
}
