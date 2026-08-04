using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize the variable with a default value.
    public int Total { get; set; } = 0;
}

public class Program
{
    public static void Main()
    {
        // Create a template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a placeholder that references the model's Total property.
        builder.Writeln("The total is: <<[model.Total]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        Document report = new Document(templatePath);

        // Create a model instance with the desired data.
        ReportModel model = new ReportModel
        {
            Total = 123 // Example value; replace with actual logic as needed.
        };

        // Build the report using the model as the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the final document.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
