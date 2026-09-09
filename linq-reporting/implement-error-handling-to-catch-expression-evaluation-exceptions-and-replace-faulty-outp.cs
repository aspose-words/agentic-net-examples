using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Simple data fields.
        builder.Writeln("Customer: <<[model.CustomerName]>>");
        builder.Writeln("Items:");

        // Loop over the collection.
        builder.Writeln("<<foreach [item in model.Items]>>");
        // Faulty expression (division by zero) – will generate an inline error message.
        builder.Writeln(" - <<[item.Index]>>: <<[item.Name]>> - Price: <<[item.Price]>> - Faulty: <<[item.Price / 0]>>");
        builder.Writeln("<</foreach>>");

        // -----------------------------------------------------------------
        // 2. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            CustomerName = "Acme Corp",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Widget", Price = 9.99 },
                new Item { Index = 2, Name = "Gadget", Price = 19.99 }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report with inline error messages enabled.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Returns false because the template contains an expression error.
        bool success = engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 4. If errors occurred, replace the inline error text with a placeholder.
        // -----------------------------------------------------------------
        if (!success)
        {
            // The engine inserts messages that contain the word "Error".
            // Use FindReplaceOptions to perform a case‑insensitive replace.
            FindReplaceOptions replaceOptions = new FindReplaceOptions
            {
                MatchCase = false // ignore case
            };

            doc.Range.Replace("Error", "[Error]", replaceOptions);
        }

        // -----------------------------------------------------------------
        // 5. Save the resulting document.
        // -----------------------------------------------------------------
        doc.Save("Report.docx");
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, non‑nullable members are initialized).
// ---------------------------------------------------------------------
public class ReportModel
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}
