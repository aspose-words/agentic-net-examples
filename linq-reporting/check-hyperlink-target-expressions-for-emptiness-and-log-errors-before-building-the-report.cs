using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -------------------------------------------------
        // 1. Create the template document programmatically
        // -------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Items Report");
        builder.Writeln("<<foreach [item in Items]>>");
        // Show the link only when the URL is not empty
        builder.Writeln("<<if [item.Url]>><<link [item.Url] [item.Name]>><</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template back before building the report
        // -------------------------------------------------
        var loadedTemplate = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare sample data
        // -------------------------------------------------
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Google", Url = "https://www.google.com" },
                new Item { Name = "EmptyLink", Url = "" }, // This will trigger an error
                new Item { Name = "Aspose", Url = "https://www.aspose.com" }
            }
        };

        // -------------------------------------------------
        // 4. Validate hyperlink targets and log errors
        // -------------------------------------------------
        foreach (var item in model.Items)
        {
            if (string.IsNullOrWhiteSpace(item.Url))
            {
                Console.WriteLine($"Error: Item '{item.Name}' has an empty hyperlink target.");
            }
        }

        // -------------------------------------------------
        // 5. Build the report using LINQ Reporting Engine
        // -------------------------------------------------
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;
        bool success = engine.BuildReport(loadedTemplate, model, "model");

        // -------------------------------------------------
        // 6. Save the generated report
        // -------------------------------------------------
        const string outputPath = "report.docx";
        loadedTemplate.Save(outputPath);

        Console.WriteLine(success
            ? $"Report generated successfully: {outputPath}"
            : $"Report generation completed with errors. See the document: {outputPath}");
    }
}

// -------------------------------------------------
// Data model classes
// -------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Url { get; set; } = string.Empty;
}
