using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingToPng
{
    // Simple data class used as the data source for the LINQ reporting engine.
    public class ReportData
    {
        public string Value { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a blank Word document.
            Document doc = new Document();

            // 2. Insert a template expression that the ReportingEngine will evaluate.
            //    The expression uses the data source name "data" and accesses its "Value" property.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Result: <<[data.Value]>>");

            // 3. Build the report by populating the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            ReportData source = new ReportData { Value = "Hello, Aspose!" };
            // The third parameter is the name used inside the template to reference the data source.
            engine.BuildReport(doc, source, "data");

            // 4. Render the populated document to a PNG image.
            //    ImageSaveOptions allows us to specify rendering options such as DPI.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render the first page (index 0). Adjust if the document has multiple pages.
                PageSet = new PageSet(0),
                // Optional: increase resolution for higher quality output.
                Resolution = 300
            };

            // 5. Save the rendered page as a PNG file.
            doc.Save("ReportResult.png", pngOptions);
        }
    }
}
