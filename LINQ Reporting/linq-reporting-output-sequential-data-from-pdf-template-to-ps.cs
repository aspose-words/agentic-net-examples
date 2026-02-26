using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the PDF template that contains LINQ Reporting tags.
        Document doc = new Document("Template.pdf");

        // Prepare a sequential data source that the template will consume.
        var data = new List<Record>
        {
            new Record { Id = 1, Name = "Alice", Amount = 123.45 },
            new Record { Id = 2, Name = "Bob",   Amount = 678.90 },
            new Record { Id = 3, Name = "Carol", Amount = 234.56 }
        };

        // Build the report using the ReportingEngine.
        // The template can reference fields like <<[data.Id]>>, <<[data.Name]>>, <<[data.Amount]>>.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "data");

        // Save the populated document as a PostScript file.
        PsSaveOptions psOptions = new PsSaveOptions
        {
            // SaveFormat is automatically set for PsSaveOptions, but can be explicit.
            SaveFormat = SaveFormat.Ps
        };
        doc.Save("Report.ps", psOptions);
    }

    // Simple POCO class used as the data source for the template.
    public class Record
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public double Amount { get; set; }
    }
}
