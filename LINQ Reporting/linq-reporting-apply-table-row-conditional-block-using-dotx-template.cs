using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOTX template that contains a table with a conditional block.
        // Example template syntax inside the table row:
        // <<if [data.Items[i].IsActive]>><<[data.Items[i].Name]>> <<[data.Items[i].Quantity]>> <<endif>>
        Document template = new Document("Template.dotx");

        // Prepare the data source that will be referenced in the template.
        var reportData = new ReportData
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Quantity = 10, IsActive = true  },
                new Item { Name = "Banana", Quantity = 0,  IsActive = false },
                new Item { Name = "Cherry", Quantity = 25, IsActive = true  }
            }
        };

        // Configure the ReportingEngine.
        // RemoveEmptyParagraphs cleans up rows that become empty after the conditional block is removed.
        // AllowMissingMembers prevents exceptions if a template references a member that does not exist.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs |
                      ReportBuildOptions.AllowMissingMembers
        };

        // Build the report. The third argument is the name used in the template to reference the data source.
        engine.BuildReport(template, reportData, "data");

        // Save the populated document.
        template.Save("Result.docx");
    }

    // Root object for the template.
    public class ReportData
    {
        public List<Item> Items { get; set; }
    }

    // Individual row data.
    public class Item
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
        public bool IsActive { get; set; }
    }
}
