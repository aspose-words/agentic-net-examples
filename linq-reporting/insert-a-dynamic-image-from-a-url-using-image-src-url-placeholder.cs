using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model that provides the image URL.
    public class ReportModel
    {
        // URL of the image to be inserted into the report.
        public string ImageUrl { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create the template document with a textbox that contains the
            //    LINQ Reporting image tag.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a textbox that will act as the image container.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
            // Move the cursor inside the textbox.
            builder.MoveTo(textBox.FirstParagraph);
            // Write the image tag. The expression refers to the model's ImageUrl property.
            builder.Write("<<image [model.ImageUrl] -fitSize>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source (model) with a real image URL.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                // Example public image URL. Replace with any reachable image if needed.
                ImageUrl = "https://www.w3.org/Icons/WWW/w3c_home_nb.png"
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The root object name must match the tag prefix ("model").
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            loadedTemplate.Save(outputPath);
        }
    }
}
