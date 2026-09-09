using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}

public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template document.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Prepare sample data.
        // -----------------------------------------------------------------
        Order order = new Order { CustomerName = "John Doe" };

        // Use a foreach loop with 'var' to let the compiler infer the type.
        foreach (var i in Enumerable.Range(1, 5))
        {
            order.Items.Add(new Item { Index = i, Name = $"Product {i}" });
        }

        // -----------------------------------------------------------------
        // Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
