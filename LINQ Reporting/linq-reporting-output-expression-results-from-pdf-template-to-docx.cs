using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the PDF template that contains LINQ Reporting tags.
        Document pdfTemplate = new Document("Template.pdf");

        // Prepare a simple data source that will be referenced from the template.
        var data = new ReportData
        {
            Title = "Quarterly Sales Report",
            Items = new[]
            {
                new Item { Name = "Apple",  Quantity = 120, Price = 1.25 },
                new Item { Name = "Banana", Quantity =  85, Price = 0.80 },
                new Item { Name = "Cherry", Quantity =  60, Price = 2.10 }
            }
        };

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Populate the template with the data source.
        // The third argument ("report") is the name used to reference the data source inside the template.
        engine.BuildReport(pdfTemplate, data, "report");

        // Save the resulting document as DOCX.
        pdfTemplate.Save("Result.docx", SaveFormat.Docx);
    }

    // Data model used by the template.
    public class ReportData
    {
        public string Title { get; set; }
        public Item[] Items { get; set; }
    }

    public class Item
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
        public double Price { get; set; }
    }
}
