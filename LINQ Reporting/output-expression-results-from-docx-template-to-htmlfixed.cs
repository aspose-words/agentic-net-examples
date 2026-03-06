using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains expression tags.
        Document doc = new Document("Template.docx");

        // Example data source – replace with your actual data object.
        var dataSource = new
        {
            Title = "Quarterly Report",
            Total = 98765.43,
            Date = DateTime.Now
        };

        // Populate the template with the data source.
        // The third argument is the name used to reference the data source in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "ds");

        // Configure options for saving to fixed HTML.
        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
        {
            // Explicitly set the format to HtmlFixed.
            SaveFormat = SaveFormat.HtmlFixed,

            // Export form fields as interactive HTML input elements (optional).
            ExportFormFields = true,

            // Do not embed images; they will be saved as separate files.
            ExportEmbeddedImages = false,

            // Optimize the output by removing redundant canvases and merging glyphs.
            OptimizeOutput = true
        };

        // Save the populated document as fixed HTML.
        doc.Save("Result.html", saveOptions);
    }
}
