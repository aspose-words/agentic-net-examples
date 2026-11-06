using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public class ReportExample
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a foreach block that iterates over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Explicitly cast the custom type to string inside the expression.
        builder.Writeln("<<[(string)item.Value]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a local file.
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template back for report generation.
        Document loadedTemplate = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Value = new CustomValue(10) },
                new Item { Value = new CustomValue(20) },
                new Item { Value = new CustomValue(30) }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        const string outputPath = "output.docx";
        loadedTemplate.Save(outputPath);
    }
}

// Wrapper class for the root data source.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Sample item class containing the custom type.
public class Item
{
    public CustomValue Value { get; set; } = new();
}

// Custom type with an explicit conversion operator to string.
public struct CustomValue
{
    public int Number { get; set; }

    public CustomValue(int number) => Number = number;

    // Explicit conversion to string.
    public static explicit operator string(CustomValue value) => $"CustomValue: {value.Number}";
}
