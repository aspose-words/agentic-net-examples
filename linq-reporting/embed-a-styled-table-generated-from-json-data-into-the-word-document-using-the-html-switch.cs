using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    public string HtmlTable { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register the code page provider required by Aspose.Words for certain encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON data.
        string json = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"",   ""Age"": 25 },
            { ""Name"": ""Charlie"", ""Age"": 28 }
        ]";

        // Deserialize JSON into a list of Person objects.
        List<Person> persons = JsonConvert.DeserializeObject<List<Person>>(json) ?? new();

        // Build an HTML table string with simple styling.
        var sb = new StringBuilder();
        sb.Append("<table style='border:1px solid black; border-collapse:collapse;'>");
        sb.Append("<tr>");
        sb.Append("<th style='border:1px solid black; padding:5px;'>Name</th>");
        sb.Append("<th style='border:1px solid black; padding:5px;'>Age</th>");
        sb.Append("</tr>");

        foreach (var p in persons)
        {
            sb.Append("<tr>");
            sb.Append($"<td style='border:1px solid black; padding:5px;'>{WebUtility.HtmlEncode(p.Name)}</td>");
            sb.Append($"<td style='border:1px solid black; padding:5px;'>{p.Age}</td>");
            sb.Append("</tr>");
        }

        sb.Append("</table>");

        // Prepare the model for the LINQ Reporting engine.
        var model = new ReportModel { HtmlTable = sb.ToString() };

        // Create a Word template programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report generated from JSON data:");
        // Insert the HTML expression tag.
        builder.Writeln("<<[model.HtmlTable] -html>>");
        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
