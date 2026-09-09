using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

public class LinqReportingExample
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample JSON data file (customers.json) in the working directory.
        // -----------------------------------------------------------------
        string jsonPath = "customers.json";
        string jsonContent = @"[
  { ""Name"": ""John Doe"", ""Address"": ""123 Main St, Anytown"", ""Email"": ""john.doe@example.com"" },
  { ""Name"": ""Jane Smith"", ""Address"": ""456 Oak Ave, Othertown"", ""Email"": ""jane.smith@example.com"" },
  { ""Name"": ""Bob Johnson"", ""Address"": ""789 Pine Rd, Sometown"", ""Email"": ""bob.johnson@example.com"" }
]";
        File.WriteAllText(jsonPath, jsonContent);

        // -----------------------------------------------------------------
        // 2. Create a Word template programmatically and save it (template.docx).
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("Customer Report");

        // Begin foreach loop over the JSON root array named "customers".
        builder.Writeln("<<foreach [c in customers]>>");

        // Create a table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Address");
        builder.InsertCell();
        builder.Writeln("Email");
        builder.EndRow();

        // Table row bound to each customer.
        builder.InsertCell();
        builder.Writeln("<<[c.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[c.Address]>>");
        builder.InsertCell();
        builder.Writeln("<<[c.Email]>>");
        builder.EndRow();

        // End the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the JSON data using ReportingEngine.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // Build the report. The root name "customers" must match the name used in the template tags.
        engine.BuildReport(reportDoc, jsonDataSource, "customers");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = "CustomerReport.docx";
        reportDoc.Save(outputPath);
    }
}
