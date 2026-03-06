using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains Aspose.Words reporting tags.
        Document doc = new Document("Template.docx");

        // Example data source – can be any POCO, DataSet, etc.
        var dataSource = new
        {
            Name = "John Doe",
            Age = 30,
            Address = "123 Main St"
        };

        // Populate the template with the data source.
        ReportingEngine engine = new ReportingEngine();
        // The third argument ("ds") is the name used in the template to reference the data source.
        engine.BuildReport(doc, dataSource, "ds");

        // Configure image save options – render the document (first page) to PNG.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        pngOptions.Resolution = 300; // optional: high‑resolution output

        // Save the rendered result as a PNG image.
        doc.Save("Result.png", pngOptions);
    }
}
