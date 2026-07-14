using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Newtonsoft.Json; // Required by the task (not used directly)

namespace LinqReportingPdfAExample
{
    public class Program
    {
        public static void Main()
        {
            // File paths for the temporary files used in the example
            string templatePath = "Template.docx";
            string xmlDataPath = "Orders.xml";
            string outputPdfPath = "Report.pdf";

            // 1. Create the template document programmatically
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Orders Report");
            builder.Writeln($"Generated on: {DateTime.Now}");
            builder.Writeln(); // empty line

            // Begin a foreach loop over the Order elements in the XML data source
            // The data source name passed to BuildReport is "orders", so we reference it directly.
            builder.Writeln("<<foreach [order in orders]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Date: <<[order.OrderDate]>>");
            builder.Writeln("Total: $<<[order.Total]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // 2. Create a sample XML data source file
            string xmlContent =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <OrderDate>2023-01-01</OrderDate>
        <Total>123.45</Total>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <OrderDate>2023-02-15</OrderDate>
        <Total>678.90</Total>
    </Order>
</Orders>";
            File.WriteAllText(xmlDataPath, xmlContent);

            // 3. Load the template document for reporting
            Document reportDoc = new Document(templatePath);

            // 4. Create the XML data source
            XmlDataSource xmlDataSource = new XmlDataSource(xmlDataPath);

            // 5. Build the report using the ReportingEngine
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // assign via property as required
            engine.BuildReport(reportDoc, xmlDataSource, "orders");

            // 6. Save the report as a PDF/A‑1b compliant document
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b
            };
            reportDoc.Save(outputPdfPath, pdfOptions);
        }
    }
}
