using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCM template that contains LINQ Reporting syntax.
        Document template = new Document("Template.docm");

        // Prepare the data source that will be bound to the template.
        var reportData = new ReportData
        {
            Title = "Quarterly Sales Report",
            Items = new List<Item>
            {
                new Item { Name = "Product A", Quantity = 120, ShowQuantity = true },
                new Item { Name = "Product B", Quantity = 45,  ShowQuantity = false },
                new Item { Name = "Product C", Quantity = 78,  ShowQuantity = true }
            }
        };

        // Configure the ReportingEngine.
        // RemoveEmptyParagraphs ensures that rows/columns that become empty after evaluation are removed.
        // AllowMissingMembers prevents exceptions if a template references a member that is not present.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.AllowMissingMembers
        };

        // Build the report. The third argument ("data") is the name used inside the template to reference the root object.
        engine.BuildReport(template, reportData, "data");

        // Save the populated document.
        template.Save("Result.docx");
    }

    // Root object for the template.
    public class ReportData
    {
        public string Title { get; set; }
        public List<Item> Items { get; set; }
    }

    // Item that will be iterated in a table.
    public class Item
    {
        public string Name { get; set; }
        public int Quantity { get; set; }

        // This flag is used in the template to conditionally display the Quantity column.
        public bool ShowQuantity { get; set; }
    }
}
