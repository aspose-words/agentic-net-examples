using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample JSON data – an array of person objects.
        string json = @"[
            { ""Id"": 1, ""Name"": ""Alice"",   ""Age"": 30 },
            { ""Id"": 2, ""Name"": ""Bob"",     ""Age"": 25 },
            { ""Id"": 3, ""Name"": ""Charlie"", ""Age"": 35 }
        ]";

        // Create a JsonDataSource from the JSON string.
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));
        var jsonDataSource = new JsonDataSource(jsonStream);

        // Build the template document programmatically.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("People Report");
        builder.Writeln();

        // Header table (static header, not repeated).
        builder.StartTable();
        builder.InsertCell(); builder.Writeln("Id");
        builder.InsertCell(); builder.Writeln("Name");
        builder.InsertCell(); builder.Writeln("Age");
        builder.EndRow();
        builder.EndTable();

        // Data rows – repeat for each person in the JSON array.
        builder.Writeln("<<foreach [p in persons]>>");
        var dataTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("<<[p.Id]>>");
        builder.InsertCell(); builder.Writeln("<<[p.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[p.Age]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using the JSON data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "persons");

        // Save the generated report.
        doc.Save("PeopleReport.docx");
    }
}
