using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsXpsExport
{
    // Simple data model for sequential data.
    public class ReportData
    {
        public List<Item> Items { get; set; } = new List<Item>();
    }

    public class Item
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains reporting tags, e.g. <<foreach [Items]>><<[Name]>> - <<[Quantity]>>\n<</foreach>>
            Document template = new Document("Template.docx");

            // Prepare the data source.
            ReportData data = new ReportData();
            data.Items.Add(new Item { Name = "Apple", Quantity = 10 });
            data.Items.Add(new Item { Name = "Banana", Quantity = 20 });
            data.Items.Add(new Item { Name = "Cherry", Quantity = 30 });

            // Populate the template using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, data, "Data"); // "Data" is the name used in the template.

            // Create XPS save options.
            XpsSaveOptions xpsOptions = new XpsSaveOptions();
            // Example: enable high‑quality rendering.
            xpsOptions.UseHighQualityRendering = true;

            // Save the populated document as XPS.
            template.Save("Output.xps", xpsOptions);
        }
    }
}
