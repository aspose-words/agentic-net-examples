using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; }
    public bool InStock { get; set; }
    public double Price { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Load the DOC template that contains a table.
        // The template should have a foreach block for Items and an IF block for InStock, e.g.:
        // <<foreach [ds.Items]>>
        //   <<if [ds.InStock]>><<[ds.Name]>> - In stock<<endif>>
        //   <<if [!ds.InStock]>><<[ds.Name]>> - Out of stock<<endif>>
        //   <<[ds.Price]:C>>
        // <<endforeach>>
        Document template = new Document("Template.docx");

        // Prepare the data source: a list of items.
        List<Item> items = new List<Item>
        {
            new Item { Name = "Apple",  InStock = true,  Price = 1.20 },
            new Item { Name = "Banana", InStock = false, Price = 0.80 },
            new Item { Name = "Cherry", InStock = true,  Price = 2.50 }
        };

        // Wrap the list in an object so the template can reference it via a name.
        var dataSource = new { Items = items };

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report. The data source name "ds" is used in the template.
        engine.BuildReport(template, dataSource, "ds");

        // Save the populated document.
        template.Save("Report.docx");
    }
}
