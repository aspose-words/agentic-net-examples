using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // -------------------------------------------------
        // Step 1: Create the LINQ Reporting template.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a greeting line.
        builder.Writeln("Dear <<[customer.Name]>>,");
        builder.Writeln();

        // Conditional block: show promotional banner only for loyal customers.
        builder.Writeln("<<if [customer.IsLoyal]>>");
        builder.Writeln("<<[customer.PromoBanner]>>");
        builder.Writeln("<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Step 2: Load the template back for report generation.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // -------------------------------------------------
        // Step 3: Prepare sample data.
        // -------------------------------------------------
        Customer sampleCustomer = new Customer
        {
            Name = "John Doe",
            IsLoyal = true,
            PromoBanner = "Exclusive Offer: 20% Discount on your next purchase!"
        };

        // -------------------------------------------------
        // Step 4: Build the report.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, sampleCustomer, "customer");

        // -------------------------------------------------
        // Step 5: Save the generated report.
        // -------------------------------------------------
        doc.Save(outputPath);
    }
}

// Public data model for the report.
public class Customer
{
    public string Name { get; set; } = string.Empty;
    public bool IsLoyal { get; set; }
    public string PromoBanner { get; set; } = string.Empty;
}
