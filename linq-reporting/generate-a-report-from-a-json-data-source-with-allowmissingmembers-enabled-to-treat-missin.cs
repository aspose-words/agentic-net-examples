using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable code page support for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample JSON data with a missing Email field for the first person.
        const string jsonFileName = "people.json";
        var jsonContent = @"[
            { ""Name"": ""John Doe"", ""Age"": 30 },
            { ""Name"": ""Jane Smith"", ""Age"": 25, ""Email"": ""jane@example.com"" }
        ]";
        File.WriteAllText(jsonFileName, jsonContent);

        // Build a template document containing LINQ Reporting tags.
        const string templateFileName = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Email: <<[person.Email]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templateFileName);

        // Load the template for reporting.
        var reportDoc = new Document(templateFileName);

        // Load JSON data source.
        var jsonDataSource = new JsonDataSource(jsonFileName);

        // Configure the reporting engine to treat missing members as null.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "N/A";

        // Build the report. The root object name in the template is "persons".
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // Save the generated report.
        const string outputFileName = "report.docx";
        reportDoc.Save(outputFileName);
    }
}
