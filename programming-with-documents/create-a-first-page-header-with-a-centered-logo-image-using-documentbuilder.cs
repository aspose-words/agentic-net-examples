using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsHeaderExample
{
    public class Program
    {
        public static void Main()
        {
            // Create an output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Prepare a simple PNG image (1x1 pixel) as a byte array.
            // This avoids the need for System.Drawing.
            byte[] logoBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lKXKAAAAAElFTkSuQmCC");

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable a different header for the first page.
            builder.PageSetup.DifferentFirstPageHeaderFooter = true;

            // Move the cursor to the first‑page header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

            // Insert the logo image into the header from the byte array.
            Shape logoShape = builder.InsertImage(logoBytes);

            // Make the image floating (no text wrap) and place it behind the text.
            logoShape.WrapType = WrapType.None;
            logoShape.BehindText = true;

            // Position the image relative to the page and center it.
            logoShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            logoShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            logoShape.HorizontalAlignment = HorizontalAlignment.Center;
            logoShape.VerticalAlignment = VerticalAlignment.Center;

            // Return to the main document body.
            builder.MoveToSection(0);
            builder.Writeln("This is the first page of the document.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is the second page of the document.");

            // Save the document.
            string outputPath = Path.Combine(outputDir, "FirstPageHeader.docx");
            doc.Save(outputPath);
        }
    }
}
