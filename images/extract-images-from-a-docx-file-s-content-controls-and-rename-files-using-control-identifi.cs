using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (sample.png)
        // -----------------------------------------------------------------
        const string sampleImagePath = "sample.png";
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black))
        {
            graphics.DrawRectangle(pen, 10, 10, 80, 80);
        }
        bitmap.Save(sampleImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Create a DOCX with two content controls, each containing the image
        // -----------------------------------------------------------------
        const string docPath = "sample.docx";
        Document doc = new Document();

        // Helper to create a content control with an image inside
        void AddImageContentControl(string tag, string title)
        {
            // Create the StructuredDocumentTag (content control)
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
            sdt.Tag = tag;
            sdt.Title = title;

            // Create a paragraph that will hold the image
            Paragraph para = new Paragraph(doc);
            sdt.AppendChild(para);

            // Insert the content control into the document body
            doc.FirstSection.Body.AppendChild(sdt);

            // Use a DocumentBuilder positioned inside the paragraph to insert the image
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(para);
            builder.InsertImage(sampleImagePath);
        }

        // First content control
        AddImageContentControl("ImageControl1", "First Image Control");

        // Add an empty paragraph to separate the controls
        doc.FirstSection.Body.AppendChild(new Paragraph(doc));

        // Second content control
        AddImageContentControl("ImageControl2", "Second Image Control");

        // Save the document
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract images from the content controls
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection sdtNodes = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        int totalExtracted = 0;

        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            string tag = !string.IsNullOrEmpty(sdt.Tag) ? sdt.Tag : "untagged";
            int imageIndex = 0;

            NodeCollection shapeNodes = sdt.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string outFileName = $"{tag}_{imageIndex}.png";
                    shape.ImageData.Save(outFileName);
                    imageIndex++;
                    totalExtracted++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the content controls.");

        // Indicate success (non‑interactive).
        Console.WriteLine($"Extraction complete. {totalExtracted} image(s) saved.");
    }
}
