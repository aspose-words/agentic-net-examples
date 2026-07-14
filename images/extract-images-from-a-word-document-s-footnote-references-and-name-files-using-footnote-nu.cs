using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a deterministic folder for output files.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // ---------- Create a sample image ----------
        string sampleImagePath = Path.Combine(outputFolder, "sample.png");
        const int imgWidth = 100;
        const int imgHeight = 100;

        // Create a white bitmap using Aspose.Drawing.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            // Save the bitmap to a file that will be used as the image source.
            bitmap.Save(sampleImagePath);
        }

        // ---------- Build a Word document with footnotes that contain images ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some regular text.
        builder.Writeln("This document demonstrates extracting images from footnotes.");

        // First footnote with an image.
        Footnote footnote1 = builder.InsertFootnote(FootnoteType.Footnote, string.Empty);
        // Move the builder into the footnote paragraph and insert the image.
        builder.MoveTo(footnote1.FirstParagraph);
        builder.InsertImage(sampleImagePath);
        // Return the builder to the main document body before adding more content.
        builder.MoveToDocumentEnd();

        // Add more body text.
        builder.Writeln("More body text after the first footnote.");

        // Second footnote with another image (reuse the same sample image for simplicity).
        Footnote footnote2 = builder.InsertFootnote(FootnoteType.Footnote, string.Empty);
        builder.MoveTo(footnote2.FirstParagraph);
        builder.InsertImage(sampleImagePath);
        // Return to the main story again.
        builder.MoveToDocumentEnd();

        // Save the document.
        string docPath = Path.Combine(outputFolder, "FootnoteImages.docx");
        doc.Save(docPath);

        // ---------- Extract images from footnote references ----------
        // Reload the document to simulate a fresh load (optional).
        Document loadedDoc = new Document(docPath);

        // Get all footnote nodes.
        NodeCollection footnoteNodes = loadedDoc.GetChildNodes(NodeType.Footnote, true);
        int extractedImageCount = 0;

        foreach (Footnote footnote in footnoteNodes.OfType<Footnote>())
        {
            // Find all Shape nodes inside the footnote.
            NodeCollection shapeNodes = footnote.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Use the actual reference mark (the number shown in the document) for naming.
                    string referenceMark = footnote.ActualReferenceMark;
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"footnote-{referenceMark}{extension}";
                    string imagePath = Path.Combine(outputFolder, imageFileName);

                    // Save the image.
                    shape.ImageData.Save(imagePath);
                    extractedImageCount++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImageCount == 0)
            throw new InvalidOperationException("No images were extracted from footnotes.");

        // Optional: list extracted files.
        Console.WriteLine($"Extracted {extractedImageCount} image(s) to folder: {outputFolder}");
    }
}
