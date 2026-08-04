using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace WatermarkExample
{
    public class Program
    {
        // Adds a configurable text watermark to the provided Document.
        public static void AddTextWatermark(
            Document doc,
            string text,
            string fontFamily = "Arial",
            float fontSize = 36f,
            Color? color = null,
            WatermarkLayout layout = WatermarkLayout.Diagonal,
            bool isSemitransparent = false)
        {
            // Prepare watermark options.
            var options = new TextWatermarkOptions
            {
                FontFamily = fontFamily,
                FontSize = fontSize,
                Color = color ?? Color.Black,
                Layout = layout,
                IsSemitrasparent = isSemitransparent
            };

            // Apply the watermark.
            doc.Watermark.SetText(text, options);
        }

        public static void Main()
        {
            // Create a blank document.
            var doc = new Document();

            // Add some sample content so the watermark can be seen.
            var builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document.");
            builder.Writeln("The text watermark will appear behind this text.");

            // Apply a configurable text watermark.
            AddTextWatermark(
                doc,
                text: "CONFIDENTIAL",
                fontFamily: "Calibri",
                fontSize: 48f,
                color: Color.Red,
                layout: WatermarkLayout.Diagonal,
                isSemitransparent: true);

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document with the watermark.
            string outputPath = Path.Combine(outputDir, "WatermarkedDocument.docx");
            doc.Save(outputPath);
        }
    }
}
