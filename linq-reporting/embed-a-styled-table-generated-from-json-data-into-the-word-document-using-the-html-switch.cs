using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // File paths.
        const string templatePath = "template.docx";
        const string jsonPath = "data.json";
        const string outputPath = "report.docx";

        // Create a JSON file containing an HTML table.
        string jsonContent = @"{
    ""HtmlTable"": ""<table style='border-collapse:collapse;'>
        <tr>
            <th style='border:1px solid black;background:#D3D3D3;'>Name</th>
            <th style='border:1px solid black;background:#D3D3D3;'>Age</th>
        </tr>
        <tr>
            <td style='border:1px solid black;'>Alice</td>
            <td style='border:1px solid black;'>30</td>
        </tr>
        <tr>
            <td style='border:1px solid black;'>Bob</td>
            <td style='border:1px solid black;'>25</td>
        </tr>
    </table>""
}";
        File.WriteAllText(jsonPath, jsonContent);

        // Create a template document with an HTML switch tag.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Information:");
        builder.Writeln("<<[model.HtmlTable] -html>>");
        templateDoc.Save(templatePath);

        // Load the template.
        Document reportDoc = new Document(templatePath);

        // Load JSON data source.
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // Build the report using the JSON data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonData, "model");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
