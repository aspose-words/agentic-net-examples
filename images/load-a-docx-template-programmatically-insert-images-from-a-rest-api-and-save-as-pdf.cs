using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // File names used in the example
        const string templatePath = "template.docx";
        const string imagePath = "sampleImage.png";
        const string outputPdfPath = "result.pdf";

        // Clean up any previous files
        if (File.Exists(templatePath)) File.Delete(templatePath);
        if (File.Exists(imagePath)) File.Delete(imagePath);
        if (File.Exists(outputPdfPath)) File.Delete(outputPdfPath);

        // -------------------------------------------------
        // 1. Create a simple DOCX template
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("This is a template document.");
        templateBuilder.Writeln("Below the image will be inserted:");
        templateDoc.Save(templatePath);
        // No Dispose needed for DocumentBuilder or Document (they are not IDisposable)

        // -------------------------------------------------
        // 2. Create a sample image locally (simulating a REST API response)
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background with white
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a blue rectangle
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Blue))
                {
                    graphics.FillRectangle(brush, 20, 20, imgWidth - 40, imgHeight - 40);
                }
            }
            // Save the image to a deterministic file name
            bitmap.Save(imagePath);
        }

        // -------------------------------------------------
        // 3. Load the DOCX template
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 4. Insert the image into the document
        // -------------------------------------------------
        // Move to the end of the document to insert after the placeholder text
        builder.MoveToDocumentEnd();
        builder.InsertParagraph(); // Ensure we are on a new paragraph
        builder.InsertImage(imagePath);

        // -------------------------------------------------
        // 5. Save the document as PDF
        // -------------------------------------------------
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // 6. Validate that the PDF was created
        // -------------------------------------------------
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("PDF output file was not created.");
    }
}
