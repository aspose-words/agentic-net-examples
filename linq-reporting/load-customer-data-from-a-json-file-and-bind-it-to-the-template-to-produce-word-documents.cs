using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting; // JsonDataSource resides in this namespace

public class Program
{
    public static void Main()
    {
        // Register code page provider for proper encoding handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare file paths in the current working directory.
        string workDir = Directory.GetCurrentDirectory();
        string dataFile = Path.Combine(workDir, "customers.json");
        string templateFile = Path.Combine(workDir, "template.docx");
        string outputFile = Path.Combine(workDir, "CustomerReport.docx");

        // -----------------------------------------------------------------
        // 1. Create sample JSON data file.
        // -----------------------------------------------------------------
        string jsonContent = @"{
  ""Customers"": [
    { ""Name"": ""John Doe"", ""Email"": ""john.doe@example.com"" },
    { ""Name"": ""Jane Smith"", ""Email"": ""jane.smith@example.com"" }
  ]
}";
        File.WriteAllText(dataFile, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("<<foreach [c in Customers]>>");
        builder.Writeln("Name : <<[c.Name]>>");
        builder.Writeln("Email: <<[c.Email]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templateFile);

        // -----------------------------------------------------------------
        // 3. Load the template and the JSON data source.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templateFile);

        // Configure JSON loading to keep the root object so that the "Customers" collection is accessible.
        JsonDataLoadOptions jsonOptions = new JsonDataLoadOptions
        {
            AlwaysGenerateRootObject = true
        };
        JsonDataSource jsonData = new JsonDataSource(dataFile, jsonOptions);

        // -----------------------------------------------------------------
        // 4. Build the report.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // No root name is required because the template accesses members directly.
        engine.BuildReport(loadedTemplate, jsonData, "");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        loadedTemplate.Save(outputFile);
    }
}
