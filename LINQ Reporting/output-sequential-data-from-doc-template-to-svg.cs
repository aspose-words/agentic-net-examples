using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCX template.
        Document doc = new Document("Template.docx");

        // Prepare sequential data that will replace the template tags.
        var items = new List<Item>
        {
            new Item { Id = 1, Name = "First" },
            new Item { Id = 2, Name = "Second" },
            new Item { Id = 3, Name = "Third" }
        };

        // Populate the template using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, items, "items");

        // Set SVG save options: fill viewport, no page border, render text as placed glyphs.
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            FitToViewPort = true,
            ShowPageBorder = false,
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs
        };

        // Save the populated document as an SVG file.
        doc.Save("Output.svg", svgOptions);
    }

    // Simple data class used as the data source.
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
}
