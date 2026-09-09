using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for older encodings if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data with some missing fields.
        string jsonPath = "people.json";
        string jsonContent = @"[
            { ""Name"": ""John Doe"", ""Age"": 30 },
            { ""Age"": 25 },
            { ""Name"": ""Jane Smith"" }
        ]";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // Create a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Load the JSON data source.
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // Configure the reporting engine to treat missing members as null.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "N/A";

        // Build the report. The data source name "persons" matches the tag used in the template.
        engine.BuildReport(doc, jsonData, "persons");

        // Save the generated report.
        doc.Save("PeopleReport.docx");
    }
}
