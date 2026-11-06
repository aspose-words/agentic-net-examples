using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CustomerInfo
{
    public string Name { get; set; } = "John Doe";
    public string Email { get; set; } = "john.doe@example.com";
}

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string resultPath = Path.Combine(outputDir, "Result.docx");

        // Create template document with LINQ Reporting tags
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Report");
        builder.Writeln("Name : <<[model.Name]>>");
        builder.Writeln("Email: <<[model.Email]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Prepare model data
        CustomerInfo model = new CustomerInfo();

        // Configure ReportingEngine
        ReportingEngine engine = new ReportingEngine();

        // Build the report (root object is the model, root name is "model")
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save(resultPath);

        // Indicate completion
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to: {resultPath}");
    }
}
