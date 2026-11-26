using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a deterministic sample image (input.png) using Aspose.Drawing.
        // -----------------------------------------------------------------
        const string sampleImagePath = "input.png";
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Create bitmap and draw a simple rectangle.
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.FillRectangle(new SolidBrush(Color.LightBlue), 20, 20, imgWidth - 40, imgHeight - 40);
            }
            bitmap.Save(sampleImagePath);
        }

        // -----------------------------------------------------------------
        // Create a new Word document and insert a footnote that contains the image.
        // -----------------------------------------------------------------
        const string docPath = "Sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This is a paragraph with a footnote reference.");

        // Insert footnote and move builder into its first paragraph.
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");
        builder.MoveTo(footnote.FirstParagraph);
        builder.InsertImage(sampleImagePath);

        // Save the document.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // Load the document (demonstrates load/save lifecycle) and extract images from footnotes.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection footnotes = loadedDoc.GetChildNodes(NodeType.Footnote, true);

        int extractedCount = 0;
        int footnoteIndex = 1;

        foreach (Footnote fn in footnotes)
        {
            // Find all Shape nodes inside the current footnote.
            NodeCollection shapes = fn.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    // First save the embedded image as PNG (original format).
                    string tempPngPath = $"footnote-{footnoteIndex}.png";
                    shape.ImageData.Save(tempPngPath);

                    // Convert the PNG to JPEG using Aspose.Drawing.
                    string outputJpegPath = $"footnote-{footnoteIndex}.jpg";
                    using (Bitmap bmp = new Bitmap(tempPngPath))
                    {
                        bmp.Save(outputJpegPath, ImageFormat.Jpeg);
                    }

                    // Clean up the temporary PNG file.
                    File.Delete(tempPngPath);

                    extractedCount++;
                }
            }
            footnoteIndex++;
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from footnotes.");
        }

        // Optional clean‑up of the sample files.
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
