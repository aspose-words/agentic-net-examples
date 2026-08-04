using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with a DateTime property.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public DateTime CreatedDate { get; set; } = DateTime.Now;
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a LINQ Reporting tag that formats the DateTime as ISO 8601.
            // Use ToString within the expression to apply the desired format.
            builder.Writeln("Created: <<[model.CreatedDate.ToString(\"yyyy-MM-ddTHH:mm:ss\")]>>");

            // Prepare the data source.
            ReportModel model = new ReportModel
            {
                // Example date; you can set any DateTime you need.
                CreatedDate = new DateTime(2023, 5, 17, 14, 30, 45)
            };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated document.
            doc.Save("Report_Output.docx");
        }
    }
}
