using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Name]>>");
        builder.Writeln("<<if [item.Price > 10]>>");
        builder.Writeln(" - Expensive");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template (required by the lifecycle rule).
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template back (demonstrates load rule usage).
        Document doc = new Document(templatePath);

        // Validate that every opening tag has a matching closing tag.
        ValidateTags(doc);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", Price = 5 },
                new Item { Name = "Laptop", Price = 1200 },
                new Item { Name = "Book", Price = 15 }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Simple data model aligned with the template.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
    }

    // Validates that each opening tag has a corresponding closing tag.
    private static void ValidateTags(Document document)
    {
        string allText = document.GetText();

        // List of supported tags.
        var tags = new[]
        {
            "foreach", "if", "bookmark", "link",
            "textColor", "backColor", "cellMerge", "restartNum"
        };

        foreach (var tag in tags)
        {
            int openingCount = 0;
            int closingCount = 0;

            // Count opening tags (<<tag ...>>).
            int index = 0;
            while ((index = allText.IndexOf($"<<{tag}", index, StringComparison.Ordinal)) != -1)
            {
                // Ensure this is not a closing tag like <</tag>>.
                if (!allText.Substring(index).StartsWith($"<</{tag}"))
                    openingCount++;
                index += 2; // Move past '<<' to continue searching.
            }

            // Count closing tags (<</tag>>).
            index = 0;
            while ((index = allText.IndexOf($"<</{tag}", index, StringComparison.Ordinal)) != -1)
            {
                closingCount++;
                index += 3; // Move past '<</' to continue searching.
            }

            if (openingCount != closingCount)
                throw new InvalidOperationException(
                    $"Tag mismatch detected for '{tag}': {openingCount} opening tag(s) vs {closingCount} closing tag(s).");
        }
    }
}
