using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public double Total { get; set; } = 0;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Step 1: Create a template document programmatically
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Write placeholders for total and discount using LINQ Reporting tags
        builder.Writeln("Total: <<[model.Total]>>");
        builder.Writeln("Discount: <<[model.Total > 500 ? model.Total * 0.1 : 0]>>");

        // Save the template to disk
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Step 2: Load the template document (simulating a separate load step)
        var doc = new Document(templatePath);

        // Step 3: Prepare the data model
        var model = new ReportModel { Total = 750 }; // Example total > 500

        // Step 4: Build the report using ReportingEngine
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Step 5: Save the generated report
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
