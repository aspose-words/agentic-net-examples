using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CustomerInfo
{
    public string Name { get; set; } = "John Doe";
    public string Email { get; set; } = "john@example.com";
}

// Wrapper model that exposes the external type as a property.
public class ReportModel
{
    public CustomerInfo CustomerInfo { get; set; } = new CustomerInfo();
}

public class Program
{
    public static void Main()
    {
        // Ensure output folder exists.
        Directory.CreateDirectory("Output");

        // Create a template document with LINQ Reporting tags that reference the wrapper model.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Name: <<[model.CustomerInfo.Name]>>");
        builder.Writeln("Customer Email: <<[model.CustomerInfo.Email]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document loadedTemplate = new Document(templatePath);

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the wrapper model as the root object.
        ReportModel model = new ReportModel();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine("Output", "Report.docx");
        loadedTemplate.Save(outputPath);
    }
}
