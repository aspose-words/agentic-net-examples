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
        // Prepare deterministic folders
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a sample JPEG image using Aspose.Drawing
        string sampleImagePath = Path.Combine(workDir, "sample.jpg");
        CreateSampleJpeg(sampleImagePath, 200, 200);

        // 2. Build a Word document that contains a footnote with the image
        string docPath = Path.Combine(workDir, "DocumentWithFootnote.docx");
        BuildDocumentWithFootnoteImage(docPath, sampleImagePath);

        // 3. Load the document and extract images from footnotes
        ExtractImagesFromFootnotes(docPath, workDir);
    }

    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        // Create a white bitmap and save it as JPEG
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        graphics.Dispose();
        bitmap.Dispose();
    }

    private static void BuildDocumentWithFootnoteImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Main body text
        builder.Writeln("This is a paragraph with a footnote reference.");

        // Insert a footnote and obtain the Footnote object
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote with image:");

        // Move the builder into the footnote's paragraph before inserting the image
        builder.MoveTo(footnote.FirstParagraph);
        builder.InsertImage(imagePath);

        // Save the document
        doc.Save(docPath);
    }

    private static void ExtractImagesFromFootnotes(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);

        // Get all footnote nodes
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        int extractedCount = 0;
        int footnoteIndex = 0;

        foreach (Footnote footnote in footnotes)
        {
            footnoteIndex++;

            // Find all shape nodes inside the footnote (deep search)
            NodeCollection shapes = footnote.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    imageIndex++;
                    string fileName = Path.Combine(outputDir,
                        $"footnote-{footnoteIndex}-{imageIndex}.jpg");
                    shape.ImageData.Save(fileName);
                    extractedCount++;
                }
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from footnotes.");
    }
}
