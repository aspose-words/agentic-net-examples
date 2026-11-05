using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image (input.png)
        const string sampleImagePath = "sample.png";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    g.DrawRectangle(pen, 10, 10, 80, 80);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // Create a Word document with a footnote that contains the image
        const string docPath = "DocumentWithFootnote.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a paragraph with a footnote reference.");
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, string.Empty);
        builder.MoveTo(footnote.FirstParagraph);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Extract images from footnotes and save them as JPEG files
        int extractedCount = 0;
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        int footnoteIndex = 1;
        foreach (Footnote fn in footnotes)
        {
            NodeCollection shapes = fn.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string outputPath = $"footnote-{footnoteIndex}.jpg";
                    using (MemoryStream ms = new MemoryStream())
                    {
                        shape.ImageData.Save(ms);
                        ms.Position = 0;
                        using (Bitmap bmp = new Bitmap(ms))
                        {
                            bmp.Save(outputPath, ImageFormat.Jpeg);
                        }
                    }
                    extractedCount++;
                }
            }
            footnoteIndex++;
        }

        // Validate that at least one image was extracted
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from footnotes.");
        }

        // Clean up temporary files (optional)
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
