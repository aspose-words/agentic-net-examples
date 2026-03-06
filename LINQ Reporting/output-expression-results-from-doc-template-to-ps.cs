using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsPsExport
{
    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains expression tags (e.g. <<[Data.Property]>>)
            string templatePath = @"C:\Docs\Template.docx";

            // Path where the resulting PostScript file will be saved
            string outputPath = @"C:\Docs\Result.ps";

            // Load the template document
            Document doc = new Document(templatePath);

            // Create a data source object that matches the fields used in the template.
            // Replace this with your actual data source (e.g., a POCO, DataSet, etc.).
            var dataSource = new
            {
                Title = "Report Title",
                Date = DateTime.Now,
                Value = 12345.67
            };

            // Populate the template with data using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The second parameter is the data source; the third (optional) parameter can be omitted
            // because we do not need to reference the data source object itself in the template.
            engine.BuildReport(doc, dataSource);

            // Configure PostScript save options.
            PsSaveOptions psOptions = new PsSaveOptions
            {
                // Explicitly set the format to PostScript.
                SaveFormat = SaveFormat.Ps,

                // Example: enable booklet printing layout if required.
                // UseBookFoldPrintingSettings = true;
            };

            // Save the populated document as a PostScript file using the specified options.
            doc.Save(outputPath, psOptions);
        }
    }
}
