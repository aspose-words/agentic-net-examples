using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare the data model with an HTML snippet.
            var model = new ReportModel();

            // Create the template document containing the LINQ Reporting tag.
            const string templatePath = "Template.docx";
            CreateTemplate(templatePath);

            // Load the template.
            var doc = new Document(templatePath);

            // Build the report by merging the model into the template.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }

        private static void CreateTemplate(string path)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert the HTML expression tag that will embed the formatted HTML snippet.
            builder.Writeln("<<[model.HtmlSnippet] -html>>");

            doc.Save(path);
        }
    }

    public class ReportModel
    {
        public string HtmlSnippet { get; set; }

        public ReportModel()
        {
            // Sample HTML with formatting (bold, italic, color).
            HtmlSnippet = "<p style='color:blue;'><b>Bold Blue Text</b> and <i>italic</i> HTML snippet.</p>";
        }
    }
}
