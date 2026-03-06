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
            // Load the DOCX template that contains expression tags (e.g. <<[Data.Name]>>)
            Document doc = new Document("Template.docx");

            // Example data source with properties referenced by the template.
            var dataSource = new
            {
                Name = "John Doe",
                Age = 30,
                JoinedDate = DateTime.Now
            };

            // Populate the template with the data source using ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource);

            // Configure PostScript save options.
            PsSaveOptions psOptions = new PsSaveOptions
            {
                // Explicitly set the format to PostScript.
                SaveFormat = SaveFormat.Ps,
                // Optional: enable high‑quality rendering for better output.
                UseHighQualityRendering = true
            };

            // Save the populated document as a PostScript file.
            doc.Save("Result.ps", psOptions);
        }
    }
}
