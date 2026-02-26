using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the Word template that contains LINQ Reporting tags.
        Document doc = new Document("Template.docx");

        // Prepare a simple data source that the template will iterate over.
        var items = new List<Item>
        {
            new Item { Name = "Apple",  Quantity = 10 },
            new Item { Name = "Banana", Quantity = 20 },
            new Item { Name = "Cherry", Quantity = 30 }
        };

        // Populate the template using the ReportingEngine.
        // The template can reference the collection via the name "items".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, items, "items");

        // Configure SVG save options.
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            SaveFormat = SaveFormat.Svg,                     // Ensure SVG format.
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs, // Render text as curves.
            ShowPageBorder = false,                         // No page border in SVG.
            FitToViewPort = true                            // Make SVG fill the viewport.
        };

        // Save the populated document as an SVG file.
        doc.Save("Report.svg", svgOptions);
    }

    // Simple POCO class used as the data source for the template.
    public class Item
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
    }
}
