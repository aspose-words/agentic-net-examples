using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image using Aspose.Drawing.
        const string sampleImagePath = "sample.png";
        const int imgWidth = 100;
        const int imgHeight = 100;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // Create a Word document that contains a footnote with the image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This paragraph contains a footnote reference.");

        // Insert a footnote and move the builder into its paragraph.
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, string.Empty);
        builder.MoveTo(footnote.FirstParagraph);
        builder.InsertImage(sampleImagePath);

        // Save the document (optional, just to have a source file).
        const string docPath = "FootnoteImages.docx";
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // Extract images that are inside footnotes and save them as JPEG files.
        // -----------------------------------------------------------------
        int extractedCount = 0;
        int footnoteIndex = 0;

        foreach (Footnote fn in doc.GetChildNodes(NodeType.Footnote, true).OfType<Footnote>())
        {
            NodeCollection shapeNodes = fn.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string outFile = $"footnote-{footnoteIndex}.jpg";
                    shape.ImageData.Save(outFile);
                    extractedCount++;
                    footnoteIndex++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from footnotes.");

        // Clean up the temporary sample image.
        if (File.Exists(sampleImagePath))
            File.Delete(sampleImagePath);
    }
}
