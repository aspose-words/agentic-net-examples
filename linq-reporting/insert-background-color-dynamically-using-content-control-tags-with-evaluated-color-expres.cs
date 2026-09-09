using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags.
        // Iterate over the Items collection and apply a dynamic background color to each item name.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<backColor [item.ColorExpression]>><<[item.Name]>> <</backColor>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items =
            {
                new Item { Name = "Apple",  ColorExpression = "LightYellow" },
                new Item { Name = "Banana", ColorExpression = "LightGreen" },
                new Item { Name = "Cherry", ColorExpression = "#FFC0CB" } // HTML color code.
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("ReportWithBackground.docx");
    }
}

// Root data model for the report.
public class ReportModel
{
    // Initialize the collection to avoid nullable warnings.
    public List<Item> Items { get; set; } = new();
}

// Individual item displayed in the report.
public class Item
{
    public string Name { get; set; } = string.Empty;
    // The expression that evaluates to a color name, HTML code, etc.
    public string ColorExpression { get; set; } = string.Empty;
}
