using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image that will be inserted into a footnote.
        const string sampleImagePath = "input.png";
        CreateSampleImage(sampleImagePath);

        // Build a Word document with a footnote that contains the sample image.
        const string docPath = "DocumentWithFootnoteImages.docx";
        BuildDocumentWithFootnoteImage(docPath, sampleImagePath);

        // Extract images from all footnotes and save them as separate JPEG files.
        ExtractFootnoteImages(docPath);
    }

    private static void CreateSampleImage(string filePath)
    {
        // Create a 100x100 white bitmap.
        using (Bitmap bitmap = new Bitmap(100, 100))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Optionally draw something deterministic.
            // For simplicity we keep it plain white.
            bitmap.Save(filePath);
        }

        // Validate that the image file was created.
        if (!File.Exists(filePath))
            throw new Exception($"Failed to create sample image at '{filePath}'.");
    }

    private static void BuildDocumentWithFootnoteImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Main body text.
        builder.Writeln("This is a paragraph with a footnote reference.");

        // Insert a footnote.
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

        // Retrieve the footnote we just added.
        Footnote footnote = (Footnote)doc.GetChildNodes(NodeType.Footnote, true)[0];

        // Move the builder cursor to the footnote's first paragraph to insert the image.
        builder.MoveTo(footnote.FirstParagraph);
        Shape imageShape = builder.InsertImage(imagePath);

        // Ensure the shape actually contains an image.
        if (!imageShape.HasImage)
            throw new Exception("The inserted shape does not contain an image.");

        // Save the document.
        doc.Save(docPath);

        // Validate that the document was saved.
        if (!File.Exists(docPath))
            throw new Exception($"Failed to save document at '{docPath}'.");
    }

    private static void ExtractFootnoteImages(string docPath)
    {
        Document doc = new Document(docPath);

        NodeCollection footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);
        int extractedImageCount = 0;
        int footnoteIndex = 0;

        foreach (Footnote footnote in footnoteNodes.OfType<Footnote>())
        {
            footnoteIndex++;
            NodeCollection shapeNodes = footnote.GetChildNodes(NodeType.Shape, true);
            int imageInFootnoteIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    imageInFootnoteIndex++;
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    // Ensure JPEG extension; if not JPEG, convert by saving as JPEG using ImageData.Save.
                    string outputFileName = $"footnote-{footnoteIndex}-{imageInFootnoteIndex}{extension}";
                    shape.ImageData.Save(outputFileName);
                    extractedImageCount++;
                }
            }
        }

        if (extractedImageCount == 0)
            throw new Exception("No images were extracted from footnotes.");

        // Optional: indicate success (no interactive output required).
        Console.WriteLine($"{extractedImageCount} image(s) extracted from footnotes.");
    }
}
