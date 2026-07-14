using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingIncludeExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Category that determines which HTML fragment to include.
        public string Category { get; set; } = string.Empty;

        // Returns the file name of the HTML fragment based on the Category value.
        public string HtmlFilePath
        {
            get
            {
                // Simple logic: Category "A" uses fragment1.html, otherwise fragment2.html.
                return Category == "A" ? "fragment1.html" : "fragment2.html";
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare sample HTML fragments.
            // -----------------------------------------------------------------
            File.WriteAllText("fragment1.html", "<p style=\"color:blue;\">This is fragment <b>ONE</b>.</p>");
            File.WriteAllText("fragment2.html", "<p style=\"color:green;\">This is fragment <i>TWO</i>.</p>");

            // -----------------------------------------------------------------
            // 2. Create the template document programmatically.
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a line showing the selected category.
            builder.Writeln("Category: <<[model.Category]>>");

            // Insert the external HTML fragment based on the model's HtmlFilePath.
            // The -html switch tells the engine to treat the inserted content as HTML.
            builder.Writeln("<<[model.HtmlFilePath] -html>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Create a model instance with a specific category.
            ReportModel model = new ReportModel { Category = "A" };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save("Report.docx");
        }
    }
}
