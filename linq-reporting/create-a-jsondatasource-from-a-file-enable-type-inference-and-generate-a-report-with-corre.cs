using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // --------------------------------------------------------------------
        // Create sample JSON data file.
        // --------------------------------------------------------------------
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "data.json");
        string jsonContent = @"{
            ""Persons"": [
                { ""Name"": ""Alice"", ""Age"": 30, ""IsMember"": true },
                { ""Name"": ""Bob"",   ""Age"": 25, ""IsMember"": false }
            ]
        }";
        File.WriteAllText(jsonPath, jsonContent);

        // --------------------------------------------------------------------
        // Build the template document with LINQ Reporting tags.
        // --------------------------------------------------------------------
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Person Report");
        // The JSON root is a table named "Persons". When using JsonDataSource,
        // the root name supplied to BuildReport ("data") refers directly to that table,
        // so we iterate over the root collection itself.
        builder.Writeln("<<foreach [p in data]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("Member: <<[p.IsMember]>>");
        builder.Writeln("<</foreach>>");

        template.Save(templatePath);

        // --------------------------------------------------------------------
        // Load the template and generate the report.
        // --------------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // JsonDataSource performs type inference automatically.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "data");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
