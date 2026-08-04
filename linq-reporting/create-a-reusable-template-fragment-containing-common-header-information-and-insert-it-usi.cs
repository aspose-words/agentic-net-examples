using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Prepare directories
            string outputDir = "Output";
            string templateDir = "Templates";
            Directory.CreateDirectory(outputDir);
            Directory.CreateDirectory(templateDir);

            // Create reusable header fragment
            string headerPath = Path.Combine(templateDir, "header.docx");
            CreateHeaderTemplate(headerPath);

            // Create main template that includes the header fragment
            string mainTemplatePath = Path.Combine(templateDir, "main.docx");
            CreateMainTemplate(mainTemplatePath, headerPath);

            // Load the main template
            Document mainDoc = new Document(mainTemplatePath);

            // Prepare data model
            ReportModel model = new()
            {
                Title = "Monthly Report",
                Date = DateTime.Now.ToString("MMMM yyyy"),
                Body = "This is the body of the report generated using Aspose.Words LINQ Reporting."
            };

            // Build the report using the LINQ Reporting engine
            ReportingEngine engine = new();
            engine.BuildReport(mainDoc, model, "model");

            // Save the generated report
            string resultPath = Path.Combine(outputDir, "Report.docx");
            mainDoc.Save(resultPath);
        }

        // Creates a header fragment containing common header tags
        private static void CreateHeaderTemplate(string path)
        {
            Document headerDoc = new();
            DocumentBuilder builder = new(headerDoc);
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<<[model.Date]>>");
            builder.Writeln("------------------------------");
            headerDoc.Save(path);
        }

        // Creates the main template and inserts the header fragment using DocumentBuilder.InsertDocument
        private static void CreateMainTemplate(string path, string headerFilePath)
        {
            Document mainDoc = new();
            DocumentBuilder builder = new(mainDoc);

            // Load the header fragment and insert its content into the main template
            Document headerDoc = new(headerFilePath);
            builder.InsertDocument(headerDoc, ImportFormatMode.KeepSourceFormatting);

            // Add a blank paragraph after the header
            builder.Writeln();

            // Add the body placeholder
            builder.Writeln("<<[model.Body]>>");
            mainDoc.Save(path);
        }

        // Public data model aligned with the template tags
        public class ReportModel
        {
            public string Title { get; set; } = "";
            public string Date { get; set; } = "";
            public string Body { get; set; } = "";
        }
    }
}
