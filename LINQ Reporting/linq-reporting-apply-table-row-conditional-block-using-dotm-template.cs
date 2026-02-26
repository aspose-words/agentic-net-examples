using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; }
    public int Quantity { get; set; }
    public bool Show { get; set; }
}

public class ReportData
{
    public List<Item> Items { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Load the DOTM template that contains the LINQ Reporting syntax.
        Document template = new Document("Template.dotm");

        // Prepare the data source.
        var data = new ReportData
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Quantity = 10, Show = true },
                new Item { Name = "Banana", Quantity = 0,  Show = false },
                new Item { Name = "Cherry", Quantity = 5,  Show = true }
            }
        };

        // Build the report. The third argument ("data") is the name used in the template to reference the object.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.AllowMissingMembers
        };
        engine.BuildReport(template, data, "data");

        // Save the populated document.
        template.Save("Result.docx");
    }
}
