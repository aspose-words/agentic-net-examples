using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;

class ExtractFootnoteImages
{
    static void Main()
    {
        // Determine paths relative to the executable folder.
        string baseDir = AppContext.BaseDirectory;
        string inputPath = Path.Combine(baseDir, "InputDocument.docx");
        string outputFolder = Path.Combine(baseDir, "FootnoteImages");
        string outputDocPath = Path.Combine(baseDir, "ProcessedDocument.docx");

        Document doc;

        if (File.Exists(inputPath))
        {
            // Load existing document.
            doc = new Document(inputPath);
        }
        else
        {
            // Create a new document with a footnote that contains a tiny image.
            doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph with some text.
            builder.Writeln("Sample paragraph with a footnote.");

            // Insert a footnote.
            Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "1");

            // Create a 1x1 pixel PNG (base64 encoded).
            const string pngBase64 =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X" +
                "6V8AAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(pngBase64);

            // Create a shape for the image.
            Shape shape = new Shape(doc, ShapeType.Image);
            using (var ms = new MemoryStream(pngBytes))
            {
                shape.ImageData.SetImage(ms);
            }
            shape.Width = 20;
            shape.Height = 20;

            // Footnotes can only contain block-level nodes, so place the shape inside a paragraph.
            Paragraph para = new Paragraph(doc);
            para.AppendChild(shape);
            footnote.AppendChild(para);
        }

        // Ensure reference marks are up‑to‑date.
        doc.UpdateFields();

        // Create output folder.
        Directory.CreateDirectory(outputFolder);

        // Get all footnote nodes.
        NodeCollection footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);

        foreach (Footnote footnote in footnoteNodes)
        {
            string footnoteNumber = footnote.ActualReferenceMark ?? "0";

            // Find all image shapes inside the footnote.
            NodeCollection imageShapes = footnote.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in imageShapes)
            {
                if (shape.HasImage)
                {
                    string extension = shape.ImageData.ImageType.ToString().ToLower(); // e.g., "png"
                    string fileName = $"Footnote_{footnoteNumber}_{++imageIndex}.{extension}";
                    string filePath = Path.Combine(outputFolder, fileName);
                    shape.ImageData.Save(filePath);
                }
            }
        }

        // Save the (possibly modified) document.
        doc.Save(outputDocPath);
        Console.WriteLine("Extraction complete.");
    }
}
