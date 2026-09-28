using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingUnicodeExample
{
    // Data model with Unicode property names.
    public class Model
    {
        public string Имя { get; set; } = "Иван Иванов";
        public string Приветствие { get; set; } = "Привет, мир!";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for template and output documents.
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // Create the template document programmatically.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write LINQ Reporting tags using literal Unicode identifiers.
            builder.Writeln("<<[model.Приветствие]>>");
            builder.Writeln("<<[model.Имя]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            Document doc = new Document(templatePath);

            // Prepare the data model.
            Model model = new Model();

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(outputPath);
        }
    }
}
