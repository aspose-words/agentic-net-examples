using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Stock { get; set; }
    public double Price { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Write a simple heading.
        builder.Writeln("Items with Stock > 0 and Price < 100:");
        builder.Writeln();

        // Begin a foreach loop over the collection Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Conditional block: display only items meeting both criteria.
        builder.Writeln("<<if [item.Stock > 0 && item.Price < 100]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>> (Stock: <<[item.Stock]>>)");
        builder.Writeln("<</if>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template (could reuse the same instance, but following the rule to load after creation).
        var doc = new Document(templatePath);

        // 3. Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Pen", Stock = 10, Price = 2.5 },
                new Item { Name = "Notebook", Stock = 0, Price = 5.0 },
                new Item { Name = "Backpack", Stock = 5, Price = 120.0 },
                new Item { Name = "Mouse", Stock = 3, Price = 25.0 },
                new Item { Name = "Keyboard", Stock = 2, Price = 80.0 }
            }
        };

        // 4. Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
