using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template, JSON data and the generated report.
        const string templatePath = "Template.docx";
        const string jsonPath = "Data.json";
        const string reportPath = "Report.docx";

        // 1. Create a simple JSON file with two numeric properties.
        //    Example: { "Price": 19.99, "Quantity": 3 }
        string jsonContent = @"{ ""Price"": 19.99, ""Quantity"": 3 }";
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Build the Word template programmatically.
        //    The template contains LINQ Reporting tags that reference the JSON root object "order".
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Price   : <<[order.Price]>>");
        builder.Writeln("Quantity: <<[order.Quantity]>>");
        // Calculated field: Price * Quantity
        builder.Writeln("Total   : <<[order.Price * order.Quantity]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template back (required by the lifecycle rules).
        Document reportDoc = new Document(templatePath);

        // 4. Create a JsonDataSource from the JSON file.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The root object name used in the template is "order".
        engine.BuildReport(reportDoc, jsonDataSource, "order");

        // 6. Save the populated report.
        reportDoc.Save(reportPath);
    }
}
