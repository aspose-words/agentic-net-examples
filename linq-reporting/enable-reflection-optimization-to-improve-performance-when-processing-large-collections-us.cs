using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public class ReportGenerator
{
    public static void Main()
    {
        // Paths for the template and the generated report
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // 1. Create a template document with LINQ Reporting tags
        CreateTemplate(templatePath);

        // 2. Load the template document
        Document doc = new Document(templatePath);

        // 3. Enable reflection optimization (static property)
        ReportingEngine.UseReflectionOptimization = true;

        // 4. Prepare sample data
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Apple",  Price = 1.20 },
                new Item { Name = "Banana", Price = 0.80 },
                new Item { Name = "Cherry", Price = 2.50 }
            }
        };

        // 5. Build the report
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // 6. Save the generated report
        doc.Save(reportPath);
    }

    // Creates a simple Word template containing a foreach loop over Items
    private static void CreateTemplate(string filePath)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Add a title
        builder.Writeln("Product List");
        builder.Writeln("-----------------");

        // LINQ Reporting foreach tag
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Name: <<[item.Name]>>   Price: $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(filePath);
    }
}

// Root data model referenced in the template as <<[model.Items]>>
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item class used inside the collection
public class Item
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}
