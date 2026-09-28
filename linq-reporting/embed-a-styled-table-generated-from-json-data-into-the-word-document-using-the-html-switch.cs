using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for potential encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON data representing a collection of items.
        string jsonData = @"[
            { ""Id"": 1, ""Name"": ""Apple"",  ""Price"": 0.5 },
            { ""Id"": 2, ""Name"": ""Banana"", ""Price"": 0.3 },
            { ""Id"": 3, ""Name"": ""Cherry"", ""Price"": 0.8 }
        ]";

        // Deserialize JSON into a list of Item objects.
        List<Item> items = JsonConvert.DeserializeObject<List<Item>>(jsonData) ?? new();

        // Generate styled HTML table from the deserialized data.
        string tableHtml = GenerateHtmlTable(items);

        // Prepare the model that will be passed to the reporting engine.
        ReportModel model = new()
        {
            TableHtml = tableHtml
        };

        // Create a Word template programmatically.
        Document template = new();
        DocumentBuilder builder = new(template);

        // Insert a paragraph with the LINQ Reporting HTML switch.
        builder.Writeln("<<[model.TableHtml] -html>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        Document doc = new(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    // Generates an HTML table with basic styling from a list of items.
    private static string GenerateHtmlTable(List<Item> items)
    {
        StringBuilder sb = new();
        sb.AppendLine(@"<table style=""border-collapse:collapse;width:100%;font-family:Arial,Helvetica,sans-serif;"">");
        sb.AppendLine(@"  <tr>");
        sb.AppendLine(@"    <th style=""border:1px solid #000;padding:5px;background:#f2f2f2;"">Id</th>");
        sb.AppendLine(@"    <th style=""border:1px solid #000;padding:5px;background:#f2f2f2;"">Name</th>");
        sb.AppendLine(@"    <th style=""border:1px solid #000;padding:5px;background:#f2f2f2;"">Price</th>");
        sb.AppendLine(@"  </tr>");

        for (int i = 0; i < items.Count; i++)
        {
            Item item = items[i];
            string bgColor = i % 2 == 0 ? "#ffffff" : "#e9e9e9";
            sb.AppendLine(@"  <tr>");
            sb.AppendLine($@"    <td style=""border:1px solid #000;padding:5px;background:{bgColor};"">{item.Id}</td>");
            sb.AppendLine($@"    <td style=""border:1px solid #000;padding:5px;background:{bgColor};"">{System.Net.WebUtility.HtmlEncode(item.Name)}</td>");
            sb.AppendLine($@"    <td style=""border:1px solid #000;padding:5px;background:{bgColor};"">{item.Price:C}</td>");
            sb.AppendLine(@"  </tr>");
        }

        sb.AppendLine(@"</table>");
        return sb.ToString();
    }
}

// Model class used as the root object for the LINQ Reporting engine.
public class ReportModel
{
    public string TableHtml { get; set; } = string.Empty;
}

// Simple data class representing an item in the JSON array.
public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}
