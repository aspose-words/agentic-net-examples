using System;
using System.Collections.Generic;
using System.Text;                     // Needed for Encoding
using Aspose.Words;
using Aspose.Words.Reporting;          // LINQ Reporting engine namespace

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data model
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  BgColor = "LightYellow" },
                new Item { Name = "Banana", BgColor = "#FFFACD" },
                new Item { Name = "Cherry", BgColor = "LightCoral" }
            }
        };

        // -----------------------------------------------------------------
        // Create the template document programmatically
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // LINQ Reporting tags
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<backColor [item.BgColor]>> <<[item.Name]>> <</backColor>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report
        // -----------------------------------------------------------------
        var templateDoc = new Document(templatePath);
        var engine = new ReportingEngine();               // Use the correct ReportingEngine class
        engine.BuildReport(templateDoc, model, "model");   // Root object name must match the tags

        // Save the final report
        templateDoc.Save("Report.docx");
    }
}

// ---------------------------------------------------------------------
// Data model classes (must be public with public properties)
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string BgColor { get; set; } = string.Empty;
}
