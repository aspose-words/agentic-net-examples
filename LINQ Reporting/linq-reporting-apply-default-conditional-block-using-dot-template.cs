using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOT template file.
        Document template = new Document("Template.dotx");

        // Prepare a simple data source that will be referenced in the template.
        var data = new ReportData
        {
            Title = "Quarterly Report",
            ShowDetails = false,
            Items = new[]
            {
                new Item { Name = "Item A", Quantity = 10 },
                new Item { Name = "Item B", Quantity = 5 }
            }
        };

        // Configure the ReportingEngine.
        // AllowMissingMembers lets the engine treat missing members as null,
        // RemoveEmptyParagraphs cleans up any paragraphs that become empty after processing.
        // MissingMemberMessage defines what to output when a member is missing.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.RemoveEmptyParagraphs,
            MissingMemberMessage = "N/A"
        };

        // Build the report. The data source is referenced in the template as "ds".
        engine.BuildReport(template, data, "ds");

        // Save the populated document.
        template.Save("Result.docx");
    }

    // Simple POCO used as the data source for the template.
    public class ReportData
    {
        public string Title { get; set; }
        public bool ShowDetails { get; set; }
        public Item[] Items { get; set; }
    }

    public class Item
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
    }
}
