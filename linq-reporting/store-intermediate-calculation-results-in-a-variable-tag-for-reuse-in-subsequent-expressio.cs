using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for legacy encodings (required by Aspose.Words in some environments).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the collection "Items".
        builder.Writeln("<<foreach [item in Items]>>");

        // Output each item's name and price.
        builder.Writeln("Item: <<[item.Name]>>  Price: <<[item.Price]>>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // After the loop, display the total price using a LINQ expression.
        builder.Writeln("Total Price: <<[Items.Sum(item => item.Price)]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data source.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Price = 1.20 },
                new Item { Name = "Banana", Price = 0.80 },
                new Item { Name = "Orange", Price = 1.50 }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}
