using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReporting
{
    // Simple data source class that holds LINQ query results.
    public class ReportData
    {
        // Collection of numbers.
        public int[] Numbers { get; set; }

        // Sum of the numbers calculated via LINQ.
        public int Sum => Numbers?.Sum() ?? 0;

        // Average of the numbers calculated via LINQ.
        public double Average => Numbers?.Average() ?? 0.0;
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains LINQ Reporting tags, e.g. <<[ds.Sum]>>.
            Document doc = new Document("Template.docx");

            // Prepare the data source.
            var data = new ReportData
            {
                Numbers = new[] { 10, 20, 30, 40, 50 }
            };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // "ds" is the name used in the template to reference the data source.
            engine.BuildReport(doc, data, "ds");

            // Save the populated document as a PNG image (first page).
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the first page (zero‑based index).
                PageSet = new PageSet(0)
            };
            doc.Save("Report.png", pngOptions);
        }
    }
}
