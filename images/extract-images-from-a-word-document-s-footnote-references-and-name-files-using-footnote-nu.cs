using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing.Common namespace

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (100x100 PNG) using Aspose.Drawing.
        // -----------------------------------------------------------------
        const string sampleImagePath = "sample.png";

        // Create bitmap and draw a simple red rectangle on a white background.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Pen pen = new Pen(Color.Red))
                {
                    graphics.DrawRectangle(pen, 10, 10, 80, 80);
                }
            }

            // Save the bitmap to a local file that will be used later.
            bitmap.Save(sampleImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Build a Word document that contains a footnote with the image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Main paragraph text.
        builder.Writeln("This paragraph contains a footnote reference.");

        // Insert a footnote and move the builder into its first paragraph.
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");
        builder.MoveTo(footnote.FirstParagraph);

        // Insert the previously created image into the footnote.
        builder.InsertImage(sampleImagePath);

        // Save the document.
        const string docPath = "FootnoteImages.docx";
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract images that reside inside footnotes.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Get all footnote nodes in the document.
        NodeCollection footnoteNodes = loadedDoc.GetChildNodes(NodeType.Footnote, true);

        int extractedImages = 0;
        int footnoteIndex = 0; // Used for deterministic file naming.

        foreach (Footnote fn in footnoteNodes)
        {
            // Find all Shape nodes (potential images) inside the current footnote.
            NodeCollection shapeNodes = fn.GetChildNodes(NodeType.Shape, true);
            int shapeIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    // Determine file extension based on the image type stored in the shape.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                    // Build a deterministic file name: footnote-{footnoteIndex}-{shapeIndex}{extension}
                    string outputFileName = $"footnote-{footnoteIndex}-{shapeIndex}{extension}";

                    // Save the image to the file system.
                    shape.ImageData.Save(outputFileName);

                    extractedImages++;
                    shapeIndex++;
                }
            }

            footnoteIndex++;
        }

        // -----------------------------------------------------------------
        // 4. Validate that at least one image was extracted.
        // -----------------------------------------------------------------
        if (extractedImages == 0)
            throw new InvalidOperationException("No images were extracted from footnotes.");
    }
}
