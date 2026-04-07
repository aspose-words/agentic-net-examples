using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // URL of the image to be inserted into the report.
        public string ImageUrl { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank Word document that will serve as the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a textbox shape – the image tag must be placed inside a textbox.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
            // Move the cursor to the first paragraph of the textbox.
            builder.MoveTo(textBox.FirstParagraph);
            // Write the LINQ Reporting image tag. The expression returns a URL string.
            builder.Write("<<image [model.ImageUrl] -fitSize>>");

            // Add a normal paragraph after the textbox for visual separation (optional).
            builder.MoveToDocumentEnd();
            builder.Writeln();

            // Prepare the data source with a publicly accessible image URL.
            ReportModel model = new ReportModel
            {
                ImageUrl = "https://www.w3.org/Icons/WWW/w3c_home_nb.png"
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated document.
            doc.Save("DynamicImageReport.docx");
        }
    }
}
