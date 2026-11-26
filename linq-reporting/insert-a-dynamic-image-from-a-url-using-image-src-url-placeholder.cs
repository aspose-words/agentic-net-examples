using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using System.Text;

namespace AsposeWordsLinqReportingExample
{
    public class ReportModel
    {
        // URL of the image to be inserted dynamically.
        public string ImageUrl { get; set; } = "https://via.placeholder.com/300x200.png?text=Sample+Image";

        // Sample title for the document.
        public string Title { get; set; } = "Dynamic Image Insertion Example";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the final report.
            string templatePath = "template.docx";
            string outputPath = "report.docx";

            // -----------------------------------------------------------------
            // Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a title placeholder.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln(); // Add an empty paragraph for spacing.

            // Insert a textbox that will hold the image.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
            builder.MoveTo(textBox.FirstParagraph);
            // Image tag with fitSize switch; the expression returns a URL string.
            builder.Write("<<image [model.ImageUrl] -fitSize>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportModel model = new ReportModel();

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(outputPath);
        }
    }
}
