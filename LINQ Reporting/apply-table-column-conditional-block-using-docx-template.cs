using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains a conditional block for a table column.
        // The template should have tags like <<if [ds.ShowPrice]>><<[ds.Price]>>><</if>> inside the column.
        Document doc = new Document("Template.docx");

        // Prepare the data source. In this example we use an anonymous object with a collection named Items.
        // Each item has a Name, a Price, and a boolean ShowPrice that controls the visibility of the Price column.
        var dataSource = new
        {
            Items = new[]
            {
                new { Name = "Apple",  Price = 1.20, ShowPrice = true  },
                new { Name = "Banana", Price = 0.80, ShowPrice = false },
                new { Name = "Carrot", Price = 0.50, ShowPrice = true  }
            }
        };

        // Build the report by merging the data source into the template.
        // The third argument ("ds") is the name used in the template to reference the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "ds");

        // Save the populated document.
        doc.Save("Result.docx");
    }
}
