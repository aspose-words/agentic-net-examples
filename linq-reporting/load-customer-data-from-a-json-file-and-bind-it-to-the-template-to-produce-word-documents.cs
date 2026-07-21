using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare file paths.
        string workDir = Directory.GetCurrentDirectory();
        string jsonPath = Path.Combine(workDir, "customers.json");
        string templatePath = Path.Combine(workDir, "Template.docx");
        string outputPath = Path.Combine(workDir, "Report.docx");

        // 1. Create sample JSON data.
        string jsonContent = @"{
  ""Customers"": [
    { ""Name"": ""John Doe"", ""Address"": ""123 Main St"", ""Email"": ""john.doe@example.com"" },
    { ""Name"": ""Jane Smith"", ""Address"": ""456 Oak Ave"", ""Email"": ""jane.smith@example.com"" },
    { ""Name"": ""Bob Johnson"", ""Address"": ""789 Pine Rd"", ""Email"": ""bob.johnson@example.com"" }
  ]
}";
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Build a template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("<<foreach [c in data.Customers]>>");
        builder.Writeln("Name: <<[c.Name]>>");
        builder.Writeln("Address: <<[c.Address]>>");
        builder.Writeln("Email: <<[c.Email]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        Document loadedTemplate = new Document(templatePath);

        // 4. Create a JsonDataSource with options that generate a root object.
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
        {
            AlwaysGenerateRootObject = true
        };
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, loadOptions);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        bool success = engine.BuildReport(loadedTemplate, jsonDataSource, "data");

        // 6. Save the generated report.
        loadedTemplate.Save(outputPath);
    }
}
