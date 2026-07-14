using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingExample
{
    // Model class that matches the JSON structure.
    public class ReportModel
    {
        public string HeaderText { get; set; } = "";
        public string FooterText { get; set; } = "";
        public string BodyContent { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary files.
            const string jsonPath = "data.json";
            const string templatePath = "Template.docx";
            const string outputPath = "ReportOutput.docx";

            // 1. Create sample JSON data.
            var sampleData = new ReportModel
            {
                HeaderText = "Custom Report Header",
                FooterText = "Page footer – generated on " + DateTime.Now.ToString("yyyy-MM-dd"),
                BodyContent = "This is the main content of the report generated using Aspose.Words LINQ Reporting Engine."
            };
            string json = JsonConvert.SerializeObject(sampleData, Formatting.Indented);
            File.WriteAllText(jsonPath, json, Encoding.UTF8);

            // 2. Build the template document programmatically.
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Header with a LINQ Reporting tag.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("<<[model.HeaderText]>>");

            // Footer with a LINQ Reporting tag.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln("<<[model.FooterText]>>");

            // Main body.
            builder.MoveToDocumentEnd();
            builder.Writeln("Report Body:");
            builder.Writeln("<<[model.BodyContent]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // 3. Load the template for report generation.
            var reportDoc = new Document(templatePath);

            // 4. Load JSON data into the model.
            string jsonFromFile = File.ReadAllText(jsonPath, Encoding.UTF8);
            var model = JsonConvert.DeserializeObject<ReportModel>(jsonFromFile) ?? new ReportModel();

            // 5. Build the report using ReportingEngine.
            var engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after tag processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };
            engine.BuildReport(reportDoc, model, "model");

            // 6. Save the final report.
            reportDoc.Save(outputPath);
        }
    }
}
