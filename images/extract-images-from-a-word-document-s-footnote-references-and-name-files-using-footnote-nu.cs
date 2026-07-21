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
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(baseDir);
        string imageDir = Path.Combine(baseDir, "Images");
        Directory.CreateDirectory(imageDir);

        // Create a deterministic sample image (100x100 white PNG).
        string sampleImagePath = Path.Combine(imageDir, "sample.png");
        CreateSampleImage(sampleImagePath, 100, 100);

        // Build a document that contains footnotes with images.
        string docPath = Path.Combine(baseDir, "FootnoteImages.docx");
        BuildDocumentWithFootnoteImages(docPath, sampleImagePath);

        // Extract images from footnotes and name them using footnote numbers.
        ExtractImagesFromFootnotes(docPath, baseDir);
    }

    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Use Aspose.Drawing to create a bitmap and fill it with white.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    private static void BuildDocumentWithFootnoteImages(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First footnote with an image.
        builder.Writeln("This is some text with a footnote reference.");
        Footnote footnote1 = builder.InsertFootnote(FootnoteType.Footnote, string.Empty);
        builder.MoveTo(footnote1.FirstParagraph);
        builder.InsertImage(imagePath);
        builder.MoveToDocumentEnd(); // Return to main text.

        // Second footnote with another image.
        builder.Writeln("More text with a second footnote.");
        Footnote footnote2 = builder.InsertFootnote(FootnoteType.Footnote, string.Empty);
        builder.MoveTo(footnote2.FirstParagraph);
        builder.InsertImage(imagePath);
        builder.MoveToDocumentEnd(); // Return to main text.

        // Save the document.
        doc.Save(docPath);
    }

    private static void ExtractImagesFromFootnotes(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);

        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        int extractedCount = 0;

        for (int i = 0; i < footnotes.Count; i++)
        {
            Footnote footnote = (Footnote)footnotes[i];
            int footnoteNumber = i + 1; // Footnote numbers are 1‑based.

            NodeCollection shapes = footnote.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string outFile = Path.Combine(outputDir, $"footnote-{footnoteNumber}{extension}");
                    shape.ImageData.Save(outFile);
                    extractedCount++;
                }
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from footnotes.");
    }
}
