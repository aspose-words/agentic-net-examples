using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting; // ReportingEngine namespace
using Aspose.Words.Reporting; // ReportBuildOptions enum

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains the LINQ Reporting syntax.
        // The template should have a table with a conditional block, e.g.:
        // <<foreach [in Items]>>
        //   <<if [ShowPrice]>><<[Price]>> <<endif>>
        // <<endforeach>>
        Document doc = new Document("Template.docx");

        // Prepare the data source. Each item has a flag (ShowPrice) that determines
        // whether the Price column should be rendered for that row.
        var items = new[]
        {
            new { Name = "Item 1", Price = 12.5, ShowPrice = true  },
            new { Name = "Item 2", Price =  8.0, ShowPrice = false },
            new { Name = "Item 3", Price = 15.0, ShowPrice = true  }
        };

        // Wrap the collection in an anonymous object so that the template can reference it.
        var dataSource = new { Items = items };

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Remove empty paragraphs that may appear after the conditional block is omitted.
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report. The third argument is the name used in the template to reference the data source.
        engine.BuildReport(doc, dataSource, "dataSource");

        // Save the populated document.
        doc.Save("Result.docx");
    }
}
