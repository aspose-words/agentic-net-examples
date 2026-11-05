using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingHyperlinkExample
{
    public class HyperlinkTarget
    {
        public string Url { get; set; } = "";
        public override string ToString() => Url;
    }

    public class Order
    {
        public string Title { get; set; } = "";
        public HyperlinkTarget Url { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider required by Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a template document with a LINQ Reporting link tag.
            string templatePath = "Template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            // Convert the Url object to string using ToString() before inserting.
            builder.Writeln("<<link [order.Url.ToString()] [order.Title]>>");
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            Document reportDoc = new Document(templatePath);

            // Prepare sample data.
            Order order = new Order
            {
                Title = "Visit Aspose",
                Url = new HyperlinkTarget { Url = "https://www.aspose.com" }
            };

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, order, "order");

            // Save the generated report.
            string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
