using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "template.docx";
        const string reportPath = "report.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // The template contains a LINQ Reporting tag placed inside a <b> markup element.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("<b><<[order.CustomerName]>></b>"); // Tag inside markup.
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and preprocess it.
        // The preprocessing scans the document structure and moves any tag that is
        // wrapped by markup tags (e.g., <b>) outside of those markup tags.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        PreprocessDocument(doc);

        // -----------------------------------------------------------------
        // 3. Prepare sample data.
        // -----------------------------------------------------------------
        Order order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Apple", Price = 1.20 },
                new Item { Name = "Banana", Price = 0.80 },
                new Item { Name = "Cherry", Price = 2.50 }
            }
        };

        // -----------------------------------------------------------------
        // 4. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(reportPath);
    }

    // Scans all paragraphs in the document and moves tags that are wrapped by
    // simple markup elements (e.g., <b>, <i>, <u>) outside of those elements.
    private static void PreprocessDocument(Document doc)
    {
        // Simple regex to capture a tag wrapped by a single pair of markup tags.
        // Example: <b><<[order.CustomerName]>></b> => <<[order.CustomerName]>><b></b>
        Regex regex = new Regex(@"<(b|i|u)>(<<\[[^\]]+\]>>)</(b|i|u)>", RegexOptions.IgnoreCase);

        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            string originalText = paragraph.GetText(); // Includes the paragraph break.
            if (!regex.IsMatch(originalText))
                continue;

            // Replace the wrapped tag with the tag placed before the markup.
            string replacedText = regex.Replace(originalText, "$2<$1></$1>");

            // Clear existing runs and write the new text.
            paragraph.RemoveAllChildren();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(paragraph);
            builder.Write(replacedText);
        }
    }
}

// ---------------------------------------------------------------------
// Data model classes used by the report.
// ---------------------------------------------------------------------
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}
