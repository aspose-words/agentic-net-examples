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
        // --------------------------------------------------------------------
        // 1. Create a sample document that contains an inline image and a DATE field.
        // --------------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Intro paragraph.
        builder.Writeln("Intro paragraph before the image.");

        // Insert a tiny 1x1 pixel PNG image from a base‑64 string.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            // InsertImage returns a Shape when the image is inline.
            Shape imageShape = builder.InsertImage(imgStream);
            if (!imageShape.HasImage)
                throw new InvalidOperationException("Failed to insert an image shape.");
        }

        // Some text after the image.
        builder.Writeln("Paragraph after the image, before the field.");

        // Insert a DATE field and update fields so the result is calculated.
        builder.InsertField(FieldType.FieldDate, true);
        sourceDoc.UpdateFields();

        // Save the source document (demonstrating the create‑save lifecycle).
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // --------------------------------------------------------------------
        // 2. Extraction: range that starts inside the image shape and ends at the field.
        // --------------------------------------------------------------------
        // Load the document (demonstrating the load lifecycle).
        Document loadedDoc = new Document(sourcePath);

        // Locate the first shape that contains an image.
        Shape startShape = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                    .OfType<Shape>()
                                    .FirstOrDefault(s => s.HasImage);
        if (startShape == null)
            throw new InvalidOperationException("No image shape found in the document.");

        // Locate the first DATE field.
        Field endField = loadedDoc.Range.Fields
                                   .FirstOrDefault(f => f.Type == FieldType.FieldDate);
        if (endField == null)
            throw new InvalidOperationException("No DATE field found in the document.");

        // Determine the paragraphs that contain the start shape and the end field.
        Paragraph startParagraph = startShape.ParentNode as Paragraph;
        Paragraph endParagraph = endField.Start.ParentNode as Paragraph;

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Unable to locate the containing paragraphs.");

        // Ensure both paragraphs belong to the same story (the main body).
        Body body = loadedDoc.FirstSection.Body;
        int startIndex = body.IndexOf(startParagraph);
        int endIndex = body.IndexOf(endParagraph);
        if (startIndex < 0 || endIndex < 0 || startIndex > endIndex)
            throw new InvalidOperationException("Invalid paragraph boundaries for extraction.");

        // --------------------------------------------------------------------
        // 3. Build a new document that will hold the extracted range.
        // --------------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Clear the default empty section/paragraph.

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to correctly import nodes from the source document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        // Import each paragraph from start to end (inclusive) into the result document.
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcParagraph = body.Paragraphs[i];
            Node importedParagraph = importer.ImportNode(srcParagraph, true);
            resultBody.AppendChild(importedParagraph);
        }

        // --------------------------------------------------------------------
        // 4. Save the extracted range.
        // --------------------------------------------------------------------
        const string resultPath = "extracted-range.docx";
        resultDoc.Save(resultPath);

        // --------------------------------------------------------------------
        // 5. Validation: ensure the result contains both an image shape and a DATE field.
        // --------------------------------------------------------------------
        bool hasImage = resultDoc.GetChildNodes(NodeType.Shape, true)
                                 .OfType<Shape>()
                                 .Any(s => s.HasImage);
        bool hasField = resultDoc.Range.Fields
                                 .Any(f => f.Type == FieldType.FieldDate);

        if (!hasImage)
            throw new InvalidOperationException("Extracted document does not contain the expected image.");
        if (!hasField)
            throw new InvalidOperationException("Extracted document does not contain the expected field.");

        // Program completes without interactive input.
    }
}
