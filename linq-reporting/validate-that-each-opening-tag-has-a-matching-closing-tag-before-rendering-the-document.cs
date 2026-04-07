using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model used by the template.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public List<Item> Items { get; set; } = new()
        {
            new Item { Name = "Apple",  Price = 1.20 },
            new Item { Name = "Banana", Price = 0.80 }
        };
    }

    public class Item
    {
        public string Name  { get; set; } = "";
        public double Price { get; set; }
    }

    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "ReportTemplate.docx";
        const string outputPath   = "ReportResult.docx";

        // 1. Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Header.
        builder.Writeln("Order report for <<[order.CustomerName]>>");
        builder.Writeln();

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("Item: <<[item.Name]>> - Price: $<<[item.Price]>>");
        // Properly close the foreach block.
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Save the template to disk.
        template.Save(templatePath);

        // 2. Load the template back (required before building the report).
        Document loadedTemplate = new Document(templatePath);

        // 3. Validate that every opening tag has a matching closing tag.
        if (!ValidateTagPairs(loadedTemplate))
        {
            Console.WriteLine("Template validation failed: mismatched tags.");
            return;
        }

        // 4. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        Order data = new Order(); // sample data
        bool success = engine.BuildReport(loadedTemplate, data, "order");

        // 5. Save the generated report.
        loadedTemplate.Save(outputPath);

        Console.WriteLine(success
            ? $"Report generated successfully: {outputPath}"
            : "Report generation completed with errors.");
    }

    // Checks that each opening tag (foreach, if, bookmark, cellMerge, restartNum)
    // has a corresponding closing tag in the correct order.
    private static bool ValidateTagPairs(Document doc)
    {
        // Extract the whole document text (including tags).
        string text = doc.GetText();

        // Regex matches tags like <<foreach ...>>, <<if ...>>, <<bookmark ...>>,
        // and closing tags like <</foreach>>, <</if>>, <</bookmark>>.
        Regex tagRegex = new Regex(
            @"<<\s*(/?)\s*(foreach|if|bookmark|cellMerge|restartNum)",
            RegexOptions.Compiled);

        Stack<string> stack = new Stack<string>();

        foreach (Match match in tagRegex.Matches(text))
        {
            bool isClosing = match.Groups[1].Value == "/";
            string tagName = match.Groups[2].Value;

            if (!isClosing)
            {
                // Opening tag – push onto the stack.
                stack.Push(tagName);
            }
            else
            {
                // Closing tag – must match the top of the stack.
                if (stack.Count == 0 || stack.Pop() != tagName)
                    return false;
            }
        }

        // All tags must be closed.
        return stack.Count == 0;
    }
}
