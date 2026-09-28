using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic file names
        const string sampleImagePath = "sample.png";
        const string docPath = "sample.docx";
        const string outputFolder = "extracted";

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // 1. Create a sample image (100x100 white PNG)
        Bitmap bitmap = new Bitmap(100, 100);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        bitmap.Save(sampleImagePath, ImageFormat.Png);
        graphics.Dispose();
        bitmap.Dispose();

        // 2. Build a DOCX with two content controls, each containing the sample image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper to create a content control with an image inside
        void InsertImageIntoContentControl(string title)
        {
            // Create the StructuredDocumentTag (content control)
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
            sdt.Title = title;

            // Create a paragraph that will hold the image
            Paragraph para = new Paragraph(doc);
            sdt.AppendChild(para);

            // Append the content control to the document body
            doc.FirstSection.Body.AppendChild(sdt);

            // Move the builder inside the paragraph of the SDT and insert the image
            builder.MoveTo(para);
            builder.InsertImage(sampleImagePath);
        }

        // First content control
        InsertImageIntoContentControl("ControlOne");

        // Add a line break between controls
        builder.Writeln();

        // Second content control
        InsertImageIntoContentControl("ControlTwo");

        // Save the document
        doc.Save(docPath);

        // 3. Load the document and extract images from content controls
        Document loadedDoc = new Document(docPath);
        NodeCollection sdtNodes = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        int extractedCount = 0;

        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            // Determine a base name for extracted images using the control's title or ID
            string baseName = !string.IsNullOrEmpty(sdt.Title) ? sdt.Title : $"Control_{sdt.Id}";

            // Find all Shape nodes inside this content control
            NodeCollection shapeNodes = sdt.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string outPath = Path.Combine(outputFolder, $"{baseName}_{imageIndex}.png");
                    shape.ImageData.Save(outPath);
                    extractedCount++;
                    imageIndex++;
                }
            }
        }

        // Validation: ensure at least one image was extracted
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the content controls.");
        }

        // Optional clean‑up (commented out)
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
