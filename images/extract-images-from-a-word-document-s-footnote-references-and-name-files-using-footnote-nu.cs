using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using Aspose.Drawing;

public class ExtractFootnoteImages
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string imagesDir = Path.Combine(baseDir, "Images");
        Directory.CreateDirectory(imagesDir);
        string outputDir = Path.Combine(baseDir, "Extracted");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image that will be inserted into footnotes.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(imagesDir, "sample.png");
        CreateSampleImage(sampleImagePath);

        // -----------------------------------------------------------------
        // 2. Build a Word document that contains footnotes with images.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First footnote with an image.
        builder.Write("This is some text with a footnote");
        Footnote footnote1 = builder.InsertFootnote(FootnoteType.Footnote, string.Empty);
        // Move the builder into the footnote to insert the image.
        builder.MoveTo(footnote1.FirstParagraph);
        builder.InsertImage(sampleImagePath);
        // Return the builder to the main story before adding more content.
        builder.MoveToDocumentEnd();

        // Second footnote with another image (same sample for simplicity).
        builder.Writeln(); // start a new paragraph in the main body
        builder.Write("Another paragraph with a second footnote");
        Footnote footnote2 = builder.InsertFootnote(FootnoteType.Footnote, string.Empty);
        builder.MoveTo(footnote2.FirstParagraph);
        builder.InsertImage(sampleImagePath);
        // Return to the main story again (optional, not needed after last footnote).
        builder.MoveToDocumentEnd();

        // Save the document.
        string docPath = Path.Combine(baseDir, "FootnoteImages.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract images from footnotes.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        int extractedCount = 0;

        // Get all footnote nodes.
        NodeCollection footnoteNodes = loadedDoc.GetChildNodes(NodeType.Footnote, true);
        int footnoteIndex = 0;

        foreach (Footnote footnote in footnoteNodes)
        {
            footnoteIndex++; // Use the order of appearance as the footnote number.

            // Find all Shape nodes inside the footnote that contain images.
            NodeCollection shapeNodes = footnote.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string outFile = Path.Combine(outputDir, $"footnote-{footnoteIndex}{extension}");
                    shape.ImageData.Save(outFile);
                    extractedCount++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from footnotes.");

        // Inform the user (no interactive prompts required).
        Console.WriteLine($"Extracted {extractedCount} image(s) to \"{outputDir}\".");
    }

    // Creates a deterministic 100x100 PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath)
    {
        const int width = 100;
        const int height = 100;

        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.LightBlue);
            // Draw a simple rectangle.
            using (Pen pen = new Pen(Color.DarkBlue, 2))
            {
                graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
            }
            bitmap.Save(filePath);
        }
    }
}
