using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDiscountExample
{
    // Data model used as the root object for the report.
    public class ReportModel
    {
        public double Total { get; set; } = 0;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a blank document and a builder to insert the template tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write the template content with LINQ Reporting tags.
            builder.Writeln("Total amount: <<[model.Total]>>");
            builder.Writeln("Discount (10% if total > 500): <<[model.Total > 500 ? model.Total * 0.1 : 0]>>");

            // Prepare sample data.
            ReportModel model = new ReportModel { Total = 750 };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated document.
            doc.Save("DiscountReport.docx");
        }
    }
}
