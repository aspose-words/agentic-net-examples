using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a simple title.
        builder.Writeln("Report");

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        builder.Writeln("Quantity: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 5 },
                new Item { Name = "Banana", Quantity = 3 },
                new Item { Name = "Orange", Quantity = 7 }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Serialize the generated report to a memory stream.
        using (MemoryStream stream = new MemoryStream())
        {
            doc.Save(stream, SaveFormat.Docx);
            stream.Position = 0; // Reset for further processing (e.g., emailing).

            // Demonstrate that the stream contains data.
            Console.WriteLine($"Report generated. Stream length: {stream.Length} bytes.");
        }
    }

    // Root data model referenced by the template.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Simple item class used in the report.
    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }
}
