using System;
using System.Collections.Generic;
using System.Drawing;                     // For Size struct
using Aspose.Words;                       // Core document classes
using Aspose.Words.Reporting;             // LINQ Reporting Engine
using Aspose.Words.Saving;                // ImageSaveOptions, SaveFormat

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains LINQ Reporting tags (e.g. <<foreach [items]>><<[Name]>> <<[/foreach]>>)
        Document template = new Document("Template.docx");

        // Prepare a sequential data source – a simple list of POCO objects.
        var items = new List<Item>
        {
            new Item { Id = 1, Name = "Alpha",   Value = 12.5 },
            new Item { Id = 2, Name = "Beta",    Value = 23.0 },
            new Item { Id = 3, Name = "Gamma",   Value = 34.75 },
            new Item { Id = 4, Name = "Delta",   Value = 45.1 }
        };

        // Build the report by populating the template with the data source.
        // The third argument is the name used inside the template to reference the collection.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, items, "items");

        // Configure image save options for TIFF output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.Resolution = 300;                         // 300 DPI for good quality
        tiffOptions.ImageSize = new Size(1240, 1754);          // Approx. A4 size at 300 DPI

        // Render each page of the populated document to a separate TIFF file.
        for (int pageIndex = 0; pageIndex < template.PageCount; pageIndex++)
        {
            tiffOptions.PageSet = new PageSet(pageIndex);    // Render only the current page
            string outFile = $"Report_Page_{pageIndex + 1}.tiff";
            template.Save(outFile, tiffOptions);             // Save using the ImageSaveOptions
        }
    }

    // Simple POCO class used as the data source for the LINQ Reporting engine.
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public double Value { get; set; }
    }
}
