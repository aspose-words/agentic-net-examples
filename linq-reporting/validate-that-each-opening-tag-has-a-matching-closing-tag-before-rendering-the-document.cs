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
        // Ensure the output folder exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a simple data model.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", Price = 5 },
                new Item { Name = "Laptop", Price = 1200 }
            }
        };

        // 2. Build the template document programmatically.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // 3. Load the template for validation and reporting.
        Document template = new Document(templatePath);

        // 4. Validate that every opening LINQ Reporting tag has a matching closing tag.
        if (!ValidateLinqReportingTags(template.GetText()))
        {
            Console.WriteLine("Tag validation failed: mismatched opening/closing tags.");
            return;
        }

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // 6. Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        template.Save(reportPath);
        Console.WriteLine($"Report generated successfully at: {reportPath}");
    }

    // Creates a template containing correctly matched LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Name]>> - Price: <<[item.Price]>>");
        builder.Writeln("<<if [item.Price > 10]>>");
        builder.Writeln(" (Expensive)");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    // Validates that each opening tag (foreach, if, bookmark) has a corresponding closing tag.
    private static bool ValidateLinqReportingTags(string documentText)
    {
        // Regex matches opening tags like <<foreach ...>> and closing tags like <</foreach>>.
        Regex tagRegex = new Regex(
            @"<<\s*(?<type>foreach|if|bookmark)[^>]*>>|<</\s*(?<type>foreach|if|bookmark)\s*>>",
            RegexOptions.IgnoreCase);

        Stack<string> stack = new Stack<string>();

        foreach (Match match in tagRegex.Matches(documentText))
        {
            // Determine if this is an opening or closing tag.
            bool isOpening = match.Value.StartsWith("<<") && !match.Value.StartsWith("<</");
            string tagType = match.Groups["type"].Value.ToLowerInvariant();

            if (isOpening)
            {
                stack.Push(tagType);
            }
            else
            {
                if (stack.Count == 0)
                    return false; // Closing tag without a matching opening tag.

                string expected = stack.Pop();
                if (expected != tagType)
                    return false; // Mismatched tag types.
            }
        }

        // All opening tags must be closed.
        return stack.Count == 0;
    }
}

// Data model used by the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item class.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}
