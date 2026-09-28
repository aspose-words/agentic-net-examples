using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class DiscountReport
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create a template document.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add static fields.
        builder.Writeln("Customer: <<[model.CustomerName]>>");
        builder.Writeln("Total: $<<[model.Total]>>");

        // Conditional block: display discount only when it is greater than zero.
        builder.Writeln("<<if [model.Discount > 0]>>Discount: $<<[model.Discount]>> <</if>>");

        // Save the template to disk.
        const string templatePath = "DiscountTemplate.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            CustomerName = "John Doe",
            Total = 120.00,
            Discount = 15.00 // Change to 0 to hide the discount line.
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "DiscountReport.docx";
        reportDoc.Save(outputPath);
    }

    // Data model used by the template.
    public class ReportModel
    {
        public string CustomerName { get; set; } = "";
        public double Total { get; set; }
        public double Discount { get; set; }
    }
}
