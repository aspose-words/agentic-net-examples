using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the DOCX template that contains LINQ Reporting tags.
        // Example tag in the template:
        // <<[model.Title]>>
        // <<foreach [item in model.Items]>>
        // - <<[item.Name]>>: <<[item.Quantity]>> pcs @ $<<[item.Price]>>
        // <</foreach>>
        string templatePath = "Template.docx";

        // Path where the generated Markdown file will be saved.
        string outputPath = "Report.md";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Prepare the data source that matches the template tags.
        var model = new ReportModel
        {
            Title = "Sales Report",
            Items = new List<ReportItem>
            {
                new ReportItem { Name = "Apple",  Quantity = 120, Price = 0.50 },
                new ReportItem { Name = "Banana", Quantity =  85, Price = 0.30 },
                new ReportItem { Name = "Cherry", Quantity =  60, Price = 1.20 }
            }
        };

        // Populate the template using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the populated document as Markdown.
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            // Optional: explicitly set the format (default is Markdown).
            SaveFormat = SaveFormat.Markdown
        };
        doc.Save(outputPath, mdOptions);
    }

    // Data model used by the template.
    public class ReportModel
    {
        public string Title { get; set; }
        public List<ReportItem> Items { get; set; }
    }

    public class ReportItem
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
        public double Price { get; set; }
    }
}
