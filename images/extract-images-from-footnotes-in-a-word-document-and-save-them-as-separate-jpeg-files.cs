using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main(string[] args)
    {
        // ------------------------------------------------------------
        // 1. Create a deterministic sample image (input.png)
        // ------------------------------------------------------------
        const string sampleImagePath = "input.png";
        const int imgWidth = 200;
        const int imgHeight = 100;

        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 3))
        {
            graphics.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
        }
        bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        graphics.Dispose();
        bitmap.Dispose();

        // ------------------------------------------------------------
        // 2. Create a Word document with a footnote that contains the image
        // ------------------------------------------------------------
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a paragraph with a footnote reference.");

        // Insert footnote
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

        // Move builder to the footnote's paragraph and insert the image
        builder.MoveTo(footnote.FirstParagraph);
        builder.InsertImage(sampleImagePath);

        // Save the document
        doc.Save(docPath);

        // ------------------------------------------------------------
        // 3. Extract images from footnotes and save as JPEG files
        // ------------------------------------------------------------
        Document loadDoc = new Document(docPath);
        NodeCollection footnoteNodes = loadDoc.GetChildNodes(NodeType.Footnote, true);
        int imageIndex = 0;

        foreach (Footnote fn in footnoteNodes)
        {
            NodeCollection shapeNodes = fn.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes)
            {
                if (!shape.HasImage)
                    continue;

                imageIndex++;
                string outputImagePath = $"footnote-{imageIndex}.jpg";

                // Convert the shape's image data to JPEG using Aspose.Drawing
                byte[] imageBytes = shape.ImageData.ImageBytes;
                using (MemoryStream ms = new MemoryStream(imageBytes))
                using (Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(ms))
                {
                    bmp.Save(outputImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                }

                // Validate that the file was created
                if (!File.Exists(outputImagePath))
                {
                    throw new InvalidOperationException($"Failed to create image file '{outputImagePath}'.");
                }
            }
        }

        // ------------------------------------------------------------
        // 4. Validate that at least one image was extracted
        // ------------------------------------------------------------
        if (imageIndex == 0)
        {
            throw new InvalidOperationException("No images were extracted from footnotes.");
        }

        // ------------------------------------------------------------
        // 5. (Optional) Clean up temporary files
        // ------------------------------------------------------------
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
