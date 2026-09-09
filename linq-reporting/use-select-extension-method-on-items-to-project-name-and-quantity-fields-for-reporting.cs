using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider required by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln();

        // Use Select to project only Name and Quantity fields.
        builder.Writeln("<<foreach [item in Items.Select(i => new { i.Name, i.Quantity })]>>");
        builder.Writeln("Name: <<[item.Name]>>, Quantity: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and prepare the data source.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 10 },
                new Item { Name = "Banana", Quantity = 20 },
                new Item { Name = "Cherry", Quantity = 15 }
            }
        };

        // -----------------------------------------------------------------
        // Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(reportPath);
    }
}
