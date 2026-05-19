using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class FootnoteImageExtractor
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the sample image and the document.
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        string docPath = Path.Combine(artifactsDir, "FootnoteImages.docx");

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image using Aspose.Drawing.
        // -----------------------------------------------------------------
        const int imgWidth = 100;
        const int imgHeight = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // White background.
                g.Clear(Aspose.Drawing.Color.White);
                // Red rectangle.
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 3))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            // Save the bitmap as PNG.
            bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Build a sample Word document that contains footnotes with images.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This document demonstrates extracting images from footnotes.");

        for (int i = 1; i <= 3; i++)
        {
            // Insert a reference marker in the main text.
            builder.Write($"Reference{i}");

            // Insert the footnote.
            Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote,
                $"Footnote {i} contents: ");

            // Move the builder into the footnote so that the image is placed inside it.
            builder.MoveTo(footnote.FirstParagraph);

            // Insert the image inside the footnote.
            Shape imgShape = builder.InsertImage(sampleImagePath);
            imgShape.WrapType = WrapType.Inline;

            // Add a line break after the image for readability.
            builder.Writeln();

            // Return the builder to the main story to continue adding references.
            builder.MoveToDocumentEnd();
        }

        // Save the document.
        doc.Save(docPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the document and extract images from each footnote.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection footnoteNodes = loadedDoc.GetChildNodes(NodeType.Footnote, true);

        int extractedCount = 0;
        int footnoteIndex = 1; // Manual counter because FootnoteNumber property does not exist.

        foreach (Footnote footnote in footnoteNodes.OfType<Footnote>())
        {
            // Find all Shape nodes inside the footnote that contain images.
            NodeCollection shapeNodes = footnote.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine file extension based on image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"footnote-{footnoteIndex}{extension}";
                    string imagePath = Path.Combine(artifactsDir, imageFileName);

                    // Save the image.
                    shape.ImageData.Save(imagePath);
                    extractedCount++;
                }
            }

            footnoteIndex++;
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from footnotes.");

        Console.WriteLine($"Extracted {extractedCount} image(s) to folder: {artifactsDir}");
    }
}
