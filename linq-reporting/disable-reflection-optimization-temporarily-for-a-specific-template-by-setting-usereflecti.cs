using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create a simple template with a LINQ Reporting tag
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>!");
        templateDoc.Save(templatePath);

        // Load the template for report generation
        Document reportDoc = new Document(templatePath);

        // Prepare the data model
        ReportModel model = new ReportModel { Name = "Aspose" };

        // Temporarily disable reflection optimization for this report
        bool originalOptimization = ReportingEngine.UseReflectionOptimization;
        ReportingEngine.UseReflectionOptimization = false;
        try
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");
        }
        finally
        {
            // Restore the original setting
            ReportingEngine.UseReflectionOptimization = originalOptimization;
        }

        // Save the generated report
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }

    // Simple public data model
    public class ReportModel
    {
        public string Name { get; set; } = "";
    }
}
