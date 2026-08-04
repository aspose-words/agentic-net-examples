using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for possible encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare output folder.
            string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputFolder);

            // Paths for the template, JSON data and final report.
            string templatePath = Path.Combine(outputFolder, "Template.docx");
            string jsonPath = Path.Combine(outputFolder, "data.json");
            string reportPath = Path.Combine(outputFolder, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create a simple template document with header, footer and body.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Header with a custom field from JSON.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("<<[model.Header]>>");

            // Footer with page information from JSON.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln("Page <<[model.PageNumber]>> of <<[model.TotalPages]>>");

            // Body content.
            builder.MoveToDocumentEnd();
            builder.Writeln("Report for <<[model.CustomerName]>>");
            builder.Writeln("Generated on <<[model.GeneratedDate]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // ---------------------------------------------------------------
            // 2. Create a JSON file that contains the data for the report.
            // ---------------------------------------------------------------
            var jsonData = new
            {
                Header = "Monthly Sales Report",
                PageNumber = "1",
                TotalPages = "5",
                CustomerName = "Acme Corp",
                GeneratedDate = DateTime.Now.ToString("yyyy-MM-dd")
            };
            string jsonString = JsonConvert.SerializeObject(jsonData, Formatting.Indented);
            File.WriteAllText(jsonPath, jsonString, Encoding.UTF8);

            // ---------------------------------------------------------------
            // 3. Load the template and bind it to the JSON data source.
            // ---------------------------------------------------------------
            Document doc = new Document(templatePath);
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after tag removal.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The root name used in the template tags is "model".
            bool success = engine.BuildReport(doc, jsonDataSource, "model");

            // ---------------------------------------------------------------
            // 4. Save the generated report.
            // ---------------------------------------------------------------
            doc.Save(reportPath);

            // Optional: indicate success (no console interaction required).
            // The application will exit automatically.
        }
    }
}
