using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define paths for the template and the generated report.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "Template.docm");
        string reportPath = Path.Combine(outputDir, "Report.docm");

        // -------------------------------------------------
        // Step 1: Create a macro-enabled template document.
        // -------------------------------------------------
        Document templateDoc = new Document();

        // Create a new VBA project and add a simple macro module.
        VbaProject templateProject = new VbaProject
        {
            Name = "TemplateProject"
        };

        VbaModule macroModule = new VbaModule
        {
            Name = "StandardModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from template\"\nEnd Sub"
        };

        templateProject.Modules.Add(macroModule);
        templateDoc.VbaProject = templateProject;

        // Save the template as a macro-enabled document.
        templateDoc.Save(templatePath, SaveFormat.Docm);

        // -------------------------------------------------
        // Step 2: Load the template and clone its VBA project.
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        VbaProject clonedProject = loadedTemplate.VbaProject.Clone();

        // -------------------------------------------------
        // Step 3: Create a new report and attach the cloned VBA project.
        // -------------------------------------------------
        Document reportDoc = new Document();

        // (Optional) Add some content to the report.
        DocumentBuilder builder = new DocumentBuilder(reportDoc);
        builder.Writeln("This is a generated report that reuses macros from the template.");

        // Assign the cloned VBA project to the new document.
        reportDoc.VbaProject = clonedProject;

        // Save the report as a macro-enabled document.
        reportDoc.Save(reportPath, SaveFormat.Docm);

        // -------------------------------------------------
        // Step 4: Simple validation – ensure the macro exists.
        // -------------------------------------------------
        Document validationDoc = new Document(reportPath);
        bool hasMacros = validationDoc.HasMacros;
        int moduleCount = validationDoc.VbaProject?.Modules?.Count ?? 0;
        string moduleName = moduleCount > 0 ? validationDoc.VbaProject.Modules[0].Name : "N/A";

        Console.WriteLine($"Report generated at: {reportPath}");
        Console.WriteLine($"Has macros: {hasMacros}");
        Console.WriteLine($"Module count: {moduleCount}");
        Console.WriteLine($"First module name: {moduleName}");
    }
}
