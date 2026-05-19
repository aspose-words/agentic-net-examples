using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using Aspose.Drawing; // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;

public class ExtractFootnoteImages
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare output folder
        // -----------------------------------------------------------------
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 2. Create a deterministic sample image (sample.png)
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, 80, 80);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 3. Build a Word document that contains a footnote with the image
        // -----------------------------------------------------------------
        string docPath = Path.Combine(outputDir, "FootnoteImages.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Main body text
        builder.Writeln("This is a paragraph with a footnote reference.");

        // Insert a footnote
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "");

        // Ensure the footnote has a paragraph to host the image
        Paragraph footnoteParagraph = footnote.FirstParagraph ?? new Paragraph(doc);
        if (footnote.FirstParagraph == null)
            footnote.AppendChild(footnoteParagraph);

        // Create a shape that holds the image and append it to the footnote paragraph
        Shape imgShape = new Shape(doc, ShapeType.Image);
        imgShape.ImageData.SetImage(sampleImagePath);
        footnoteParagraph.AppendChild(imgShape);

        // Save the document
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 4. Load the document and extract images from footnotes
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection footnotes = loadedDoc.GetChildNodes(NodeType.Footnote, true);

        int extractedCount = 0;
        int footnoteIndex = 0;

        foreach (Footnote fn in footnotes)
        {
            footnoteIndex++;
            NodeCollection shapes = fn.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    // Determine file extension; force JPEG if not already JPEG/JPG
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    if (!extension.Equals(".jpeg", StringComparison.OrdinalIgnoreCase) &&
                        !extension.Equals(".jpg", StringComparison.OrdinalIgnoreCase))
                    {
                        extension = ".jpg";
                    }

                    string imageFileName = Path.Combine(
                        outputDir,
                        $"footnote-{footnoteIndex}-{imageIndex}{extension}");

                    shape.ImageData.Save(imageFileName);
                    extractedCount++;
                    imageIndex++;
                }
            }
        }

        // -----------------------------------------------------------------
        // 5. Validate that at least one image was extracted
        // -----------------------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from footnotes.");
    }
}
