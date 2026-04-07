using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize to avoid nullable warnings.
    public string Name { get; set; } = "World";
}

public class Program
{
    public static void Main()
    {
        // Create a simple template document with a LINQ Reporting tag.
        var templatePath = "Template.docx";
        var outputPath = "Report.docx";

        // Build the template.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>!");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare the data model.
        var model = new ReportModel { Name = "Aspose.Words" };

        // Use a disposable object to create a using block.
        // Inside the block we modify the static engine setting.
        using (var dummy = new MemoryStream())
        {
            // Disable reflection optimization for this report.
            ReportingEngine.UseReflectionOptimization = false;

            // Create the reporting engine and build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");
        }

        // Save the generated report.
        doc.Save(outputPath);
    }
}
