using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the WORDML (WordprocessingML) template that contains the table and conditional block.
        // The template file should be placed in the same folder as the executable or provide a full path.
        Document template = new Document("TableTemplate.xml");

        // Prepare the data source that will be used by the LINQ Reporting Engine.
        // The template will reference the data source by the name "data".
        var reportData = new ReportData
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Quantity = 10, IsDiscounted = false },
                new Item { Name = "Banana", Quantity = 5,  IsDiscounted = true  },
                new Item { Name = "Cherry", Quantity = 12, IsDiscounted = false }
            }
        };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // Remove paragraphs that become empty after processing and allow missing members to be treated as null.
            Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.AllowMissingMembers
        };

        // Build the report. The third argument is the name used inside the template to reference the data source.
        engine.BuildReport(template, reportData, "data");

        // Save the populated document.
        template.Save("Result.docx");
    }
}

// Data source root class.
public class ReportData
{
    public List<Item> Items { get; set; }
}

// Individual row class that will be iterated in the template.
public class Item
{
    public string Name { get; set; }
    public int Quantity { get; set; }
    public bool IsDiscounted { get; set; }
}
