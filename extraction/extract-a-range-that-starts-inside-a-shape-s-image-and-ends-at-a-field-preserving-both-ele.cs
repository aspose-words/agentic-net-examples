using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample source document.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert a tiny PNG image (inline shape) and keep a reference to the shape.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imgStream);
        }

        // Add some text after the image.
        builder.Writeln("Text after image.");

        // Insert a DATE field and keep a reference to it.
        Field dateField = builder.InsertField(FieldType.FieldDate, true);
        builder.Writeln("More text after field.");

        // Save the source document (demonstrating the required save rule).
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the document (demonstrating the required load rule).
        // -------------------------------------------------
        Document loaded = new Document(sourcePath);

        // -------------------------------------------------
        // 3. Locate the image shape.
        // -------------------------------------------------
        Shape imageShape = loaded.GetChildNodes(NodeType.Shape, true)
                                 .OfType<Shape>()
                                 .FirstOrDefault(s => s.HasImage);
        if (imageShape == null)
            throw new InvalidOperationException("Image shape not found.");

        // -------------------------------------------------
        // 4. Locate the first field (the DATE field we inserted).
        // -------------------------------------------------
        Field targetField = loaded.Range.Fields.FirstOrDefault();
        if (targetField == null)
            throw new InvalidOperationException("Target field not found.");

        // -------------------------------------------------
        // 5. Determine the paragraphs that contain the start (image) and end (field).
        // -------------------------------------------------
        Paragraph startParagraph = imageShape.ParentParagraph;
        Paragraph endParagraph = targetField.Start.ParentParagraph;

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Unable to determine start or end paragraph.");

        // -------------------------------------------------
        // 6. Build a new document that will contain the extracted range.
        // -------------------------------------------------
        Document result = new Document();
        result.RemoveAllChildren(); // Ensure a clean document structure.

        Section resultSection = new Section(result);
        result.AppendChild(resultSection);

        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // -------------------------------------------------
        // 7. Import (clone) paragraphs from start to end (inclusive) preserving all inline nodes.
        //    Use NodeImporter to avoid cross‑document node errors.
        // -------------------------------------------------
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        bool copying = false;
        foreach (Paragraph para in loaded.FirstSection.Body.Paragraphs)
        {
            if (!copying && para == startParagraph)
                copying = true;

            if (copying)
            {
                // Import the paragraph (deep clone) into the destination document.
                Node importedNode = importer.ImportNode(para, true);
                resultBody.AppendChild(importedNode);
            }

            if (copying && para == endParagraph)
                break;
        }

        // -------------------------------------------------
        // 8. Save the extracted range.
        // -------------------------------------------------
        const string resultPath = "extracted.docx";
        result.Save(resultPath);

        // -------------------------------------------------
        // 9. Validate that the output file was created.
        // -------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Extraction failed: output file not created.");

        Console.WriteLine("Extraction completed successfully.");
    }
}
