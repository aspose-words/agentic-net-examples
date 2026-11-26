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
        // 1. Create a sample source document with an image and a field.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert a tiny PNG image.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imgStream);
        }

        // End the paragraph that contains the image.
        builder.Writeln();

        // Insert a MERGEFIELD field.
        builder.InsertField("MERGEFIELD MyField", "Field Result");

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the document (follow load rule).
        // -------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -------------------------------------------------
        // 3. Locate the image‑bearing shape.
        // -------------------------------------------------
        Shape imageShape = loadedDoc.GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .FirstOrDefault(s => s.HasImage);
        if (imageShape == null)
            throw new InvalidOperationException("Image shape not found.");

        // -------------------------------------------------
        // 4. Locate the first field.
        // -------------------------------------------------
        Field field = loadedDoc.Range.Fields.FirstOrDefault();
        if (field == null)
            throw new InvalidOperationException("Field not found.");

        // -------------------------------------------------
        // 5. Determine the containing paragraphs.
        // -------------------------------------------------
        Paragraph startParagraph = imageShape.ParentNode as Paragraph;
        Paragraph endParagraph = field.Start.ParentNode as Paragraph;
        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Unable to determine paragraph boundaries.");

        // -------------------------------------------------
        // 6. Build a new document that contains the range
        //    from the start paragraph (inclusive) to the end paragraph (inclusive).
        // -------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Import the start paragraph into the result document.
        Node importedStart = resultDoc.ImportNode(startParagraph, true);
        resultBody.AppendChild(importedStart);

        // If the end paragraph is different, import it as well.
        if (!ReferenceEquals(endParagraph, startParagraph))
        {
            Node importedEnd = resultDoc.ImportNode(endParagraph, true);
            resultBody.AppendChild(importedEnd);
        }

        // -------------------------------------------------
        // 7. Save the extracted content.
        // -------------------------------------------------
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // -------------------------------------------------
        // 8. Validation: ensure the extracted document contains both the image and the field.
        // -------------------------------------------------
        Document validationDoc = new Document(resultPath);
        bool hasImage = validationDoc.GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .Any(s => s.HasImage);
        bool hasField = validationDoc.Range.Fields.Count > 0;

        if (!hasImage)
            throw new InvalidOperationException("Extracted document does not contain the expected image.");
        if (!hasField)
            throw new InvalidOperationException("Extracted document does not contain the expected field.");

        // Ensure the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result file was not created.");

        // Example completed successfully.
    }
}
