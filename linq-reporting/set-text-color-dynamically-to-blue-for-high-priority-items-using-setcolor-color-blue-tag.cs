using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Task A", IsHigh = true },
                new Item { Name = "Task B", IsHigh = false },
                new Item { Name = "Task C", IsHigh = true },
                new Item { Name = "Task D", IsHigh = false }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = "Template.docx";

        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Begin a foreach loop over the collection Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // If the item has high priority, wrap its name with a blue textColor tag.
        // Otherwise output the name without coloring.
        builder.Writeln(
            "<<if [item.IsHigh]>>" +
            "<<textColor [\"Blue\"]>><<[item.Name]>> <</textColor>><</if>>" +
            "<<if [!item.IsHigh]>><<[item.Name]>> <</if>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, non‑nullable members are initialized).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public bool IsHigh { get; set; }
}
