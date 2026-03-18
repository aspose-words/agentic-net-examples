using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    // Sample data model matching the JSON structure.
    private class Item
    {
        public string Name { get; set; }
        public int Qty { get; set; }
        public double Price { get; set; }
    }

    static void Main()
    {
        // 1. Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 2. JSON data that will be transformed into a styled HTML table.
        //    In a real scenario this could be read from a file, a web service, etc.
        string json = @"
        [
            { ""Name"": ""Apple"",  ""Qty"": 10, ""Price"": 1.20 },
            { ""Name"": ""Banana"", ""Qty"": 5,  ""Price"": 0.80 },
            { ""Name"": ""Carrot"", ""Qty"": 7,  ""Price"": 0.60 }
        ]";

        // 3. Deserialize the JSON into a list of Item objects.
        List<Item> items = JsonSerializer.Deserialize<List<Item>>(json);

        // 4. Build an HTML string that represents a styled table.
        //    Inline CSS is used so that the styling is preserved when the HTML is inserted.
        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.AppendLine("<style>");
        htmlBuilder.AppendLine("  table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }");
        htmlBuilder.AppendLine("  th, td { border: 1px solid #4A90E2; padding: 8px; text-align: left; }");
        htmlBuilder.AppendLine("  th { background-color: #4A90E2; color: white; }");
        htmlBuilder.AppendLine("  tr:nth-child(even) { background-color: #F2F2F2; }");
        htmlBuilder.AppendLine("</style>");

        htmlBuilder.AppendLine("<table>");
        htmlBuilder.AppendLine("  <tr>");
        htmlBuilder.AppendLine("    <th>Name</th>");
        htmlBuilder.AppendLine("    <th>Quantity</th>");
        htmlBuilder.AppendLine("    <th>Price</th>");
        htmlBuilder.AppendLine("  </tr>");

        foreach (var item in items)
        {
            htmlBuilder.AppendLine("  <tr>");
            htmlBuilder.AppendLine($"    <td>{System.Net.WebUtility.HtmlEncode(item.Name)}</td>");
            htmlBuilder.AppendLine($"    <td>{item.Qty}</td>");
            htmlBuilder.AppendLine($"    <td>{item.Price:C}</td>");
            htmlBuilder.AppendLine("  </tr>");
        }

        htmlBuilder.AppendLine("</table>");

        string htmlTable = htmlBuilder.ToString();

        // 5. Insert the generated HTML into the document.
        //    The second overload allows us to specify that the builder's formatting
        //    should be used as the base formatting for the imported HTML.
        builder.InsertHtml(htmlTable, useBuilderFormatting: true);

        // 6. Save the document to a file.
        doc.Save("StyledTableFromJson.docx", SaveFormat.Docx);
    }
}
