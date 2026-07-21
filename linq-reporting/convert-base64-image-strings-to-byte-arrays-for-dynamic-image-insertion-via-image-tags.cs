using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing; // Required for ShapeType

namespace AsposeWordsLinqReportingExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Base64‑encoded PNG image (1×1 pixel).
        public string ImageBase64 { get; set; } = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";

        // Byte array obtained from the Base64 string – this is the value the image tag will consume.
        public byte[] ImageBytes => Convert.FromBase64String(ImageBase64);
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a textbox that will contain the image tag.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
            builder.MoveTo(textBox.FirstParagraph);

            // LINQ Reporting image tag – the expression returns a byte[].
            builder.Write("<<image [model.ImageBytes] -fitSize>>");

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, model, "model");

            // Save the generated document.
            template.Save("ReportOutput.docx");
        }
    }
}
