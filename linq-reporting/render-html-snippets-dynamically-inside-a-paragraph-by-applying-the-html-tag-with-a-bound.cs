using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace AsposeWordsLinqReportingHtmlExample
{
    // Data model with an HTML snippet property.
    public class ReportModel
    {
        // Initialize with sample HTML content.
        public string HtmlSnippet { get; set; } = "<b>Bold Text</b> and <i>Italic Text</i>";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for possible encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file paths.
            string templatePath = "template.docx";
            string outputPath = "output.docx";

            // -------------------------------------------------
            // Step 1: Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a paragraph that will contain the dynamic HTML.
            builder.Writeln("Report:");
            // Insert the LINQ Reporting HTML tag bound to the model's HtmlSnippet property.
            builder.Writeln("<<html [model.HtmlSnippet]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Step 2: Load the template for report generation.
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            // -------------------------------------------------
            // Step 3: Prepare the data model.
            // -------------------------------------------------
            ReportModel model = new ReportModel();

            // -------------------------------------------------
            // Step 4: Build the report using the LINQ Reporting engine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // Step 5: Save the generated report.
            // -------------------------------------------------
            doc.Save(outputPath);

            // Optional: indicate completion.
            Console.WriteLine($"Report generated and saved to '{outputPath}'.");
        }
    }
}
