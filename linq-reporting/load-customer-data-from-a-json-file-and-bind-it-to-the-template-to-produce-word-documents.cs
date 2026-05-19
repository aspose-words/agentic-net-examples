using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory
        string jsonPath = Path.Combine(Environment.CurrentDirectory, "customers.json");
        string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create sample JSON data (array of customer objects)
        // -----------------------------------------------------------------
        string jsonContent = @"[
            { ""Name"": ""John Doe"", ""Address"": ""123 Main St, Springfield"" },
            { ""Name"": ""Jane Smith"", ""Address"": ""456 Oak Ave, Metropolis"" },
            { ""Name"": ""Bob Johnson"", ""Address"": ""789 Pine Rd, Gotham"" }
        ]";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Create a Word template with LINQ Reporting tags
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("----------------");
        // Begin foreach loop over the JSON root array named 'customers'
        builder.Writeln("<<foreach [c in customers]>>");
        builder.Writeln("Name: <<[c.Name]>>");
        builder.Writeln("Address: <<[c.Address]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template for report generation
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Create a JsonDataSource from the JSON file
        // -----------------------------------------------------------------
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // -----------------------------------------------------------------
        // 5. Build the report using ReportingEngine
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "customers");

        // -----------------------------------------------------------------
        // 6. Save the generated report
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
