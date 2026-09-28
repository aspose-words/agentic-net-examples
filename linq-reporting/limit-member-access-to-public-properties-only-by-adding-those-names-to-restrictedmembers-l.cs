using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Define file paths
        string templatePath = "ReportTemplate.docx";
        string outputPath = "ReportResult.docx";

        // -------------------------------------------------
        // Create the template document with LINQ tags
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age: <<[model.Age]>>");
        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for reporting
        // -------------------------------------------------
        var doc = new Document(templatePath);

        // -------------------------------------------------
        // Prepare the data model
        // -------------------------------------------------
        var model = new Customer
        {
            Name = "John Doe",
            Age = 30
        };

        // -------------------------------------------------
        // Configure the reporting engine
        // -------------------------------------------------
        var engine = new ReportingEngine();

        // Aspose.Words ReportingEngine accesses only public members by default.
        // If a version supports a RestrictedMembers collection, you could add:
        // engine.RestrictedMembers.Add(nameof(Customer.Name));
        // engine.RestrictedMembers.Add(nameof(Customer.Age));
        // The above lines are omitted because the property is not available in this version.

        // Build the report
        engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // Save the generated report
        // -------------------------------------------------
        doc.Save(outputPath);
    }
}

// Public data model with public properties
public class Customer
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }

    // Private member (not accessible from the template)
    private string PrivateInfo { get; set; } = "Secret";
}
