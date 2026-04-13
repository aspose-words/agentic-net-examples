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
        // Define output folder and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a deterministic sample image using Aspose.Drawing.
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        const int imgWidth = 100;
        const int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
            }
            bitmap.Save(sampleImagePath);
        }

        // Build a document with footnotes that contain the sample image.
        string docPath = Path.Combine(outputDir, "FootnoteImages.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First footnote with an image.
        builder.Write("Paragraph with first footnote reference.");
        Footnote footnote1 = builder.InsertFootnote(FootnoteType.Footnote, "First footnote text.");
        builder.MoveTo(footnote1.FirstParagraph);
        builder.InsertImage(sampleImagePath);
        builder.MoveToDocumentEnd();

        // Second footnote with an image.
        builder.Writeln();
        builder.Write("Paragraph with second footnote reference.");
        Footnote footnote2 = builder.InsertFootnote(FootnoteType.Footnote, "Second footnote text.");
        builder.MoveTo(footnote2.FirstParagraph);
        builder.InsertImage(sampleImagePath);
        builder.MoveToDocumentEnd();

        // Save the document.
        doc.Save(docPath);

        // Load the document (optional, can reuse the same instance).
        Document loadedDoc = new Document(docPath);

        int extractedCount = 0;

        // Iterate all footnote nodes.
        NodeCollection footnoteNodes = loadedDoc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote footnote in footnoteNodes)
        {
            // Determine a deterministic footnote identifier.
            string footnoteId = footnote.ActualReferenceMark; // e.g., "1", "2", ...

            // Find all shape nodes inside the footnote.
            NodeCollection shapeNodes = footnote.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    // Determine file extension based on image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"footnote-{footnoteId}{extension}";
                    string imagePath = Path.Combine(outputDir, imageFileName);

                    // Save the image.
                    shape.ImageData.Save(imagePath);
                    if (!File.Exists(imagePath))
                        throw new InvalidOperationException($"Failed to save image for footnote {footnoteId}.");

                    extractedCount++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from footnotes.");

        // Optionally, inform the user (no interactive prompts required).
        Console.WriteLine($"Extracted {extractedCount} image(s) to folder: {outputDir}");
    }
}
