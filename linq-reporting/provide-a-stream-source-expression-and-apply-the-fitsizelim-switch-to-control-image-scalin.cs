using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model that provides an image as a Stream.
    public class ReportModel
    {
        // The image stream will be supplied by the caller.
        public Stream ImageStream { get; set; } = Stream.Null;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare a tiny PNG image (1x1 pixel, fully transparent) as a byte array.
            //    The image is stored as a Base64 string to avoid external files.
            // -----------------------------------------------------------------
            const string base64Png =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE" +
                "AwJ/lbVYAAAAAElFTkSuQmCC";
            byte[] imageBytes = Convert.FromBase64String(base64Png);

            // Create a MemoryStream that will be used by the template.
            var imageStream = new MemoryStream(imageBytes);

            // -----------------------------------------------------------------
            // 2. Build the data model and assign the stream.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                ImageStream = imageStream
            };

            // -----------------------------------------------------------------
            // 3. Create the template document programmatically.
            //    The image tag must be placed inside a textbox (valid container).
            //    The -fitSizeLim switch limits the image size to the shape bounds.
            // -----------------------------------------------------------------
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert a textbox that will hold the image.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
            builder.MoveTo(textBox.FirstParagraph);

            // LINQ Reporting tag: image from a Stream with -fitSizeLim switch.
            builder.Write("<<image [model.ImageStream] -fitSizeLim>>");

            // -----------------------------------------------------------------
            // 4. Reset the stream position before the engine consumes it.
            // -----------------------------------------------------------------
            model.ImageStream.Position = 0;

            // -----------------------------------------------------------------
            // 5. Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 6. Save the resulting document.
            // -----------------------------------------------------------------
            const string outputPath = "Report_Output.docx";
            doc.Save(outputPath);

            Console.WriteLine($"Report generated and saved to '{outputPath}'.");
        }
    }
}
