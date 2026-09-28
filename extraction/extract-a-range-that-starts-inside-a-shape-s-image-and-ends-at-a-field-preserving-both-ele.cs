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
        // ------------------------------------------------------------
        // 1. Create a sample source document containing an image and a field.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert a tiny PNG image from a base‑64 string.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imgStream);
        }

        // Add some text between the image and the field.
        builder.Writeln("Some text between image and field.");

        // Insert a DATE field on its own paragraph.
        Field dateField = builder.InsertField(FieldType.FieldDate, true);
        builder.Writeln();

        // Save the source document locally.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // ------------------------------------------------------------
        // 2. Load the document for extraction.
        // ------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Locate the first shape that actually contains an image.
        Shape imageShape = loadedDoc.GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .FirstOrDefault(s => s.HasImage);
        if (imageShape == null)
            throw new InvalidOperationException("Image shape not found.");

        // Locate the first field (the DATE field we inserted).
        Field targetField = loadedDoc.Range.Fields.FirstOrDefault();
        if (targetField == null)
            throw new InvalidOperationException("Target field not found.");

        // The field resides inside a paragraph – we will extract that whole paragraph.
        Paragraph fieldParagraph = targetField.Start.ParentNode as Paragraph;
        if (fieldParagraph == null)
            throw new InvalidOperationException("Field paragraph not found.");

        // ------------------------------------------------------------
        // 3. Build a new document that will hold the extracted range.
        // ------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // ------------------------------------------------------------
        // 4. Import the image shape.
        //    Shapes are inline nodes and must be placed inside a paragraph.
        // ------------------------------------------------------------
        Paragraph shapeParagraph = new Paragraph(resultDoc);
        NodeImporter shapeImporter = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedShape = shapeImporter.ImportNode(imageShape, true);
        shapeParagraph.AppendChild(importedShape);
        resultBody.AppendChild(shapeParagraph);

        // ------------------------------------------------------------
        // 5. Import the paragraph that contains the field.
        // ------------------------------------------------------------
        NodeImporter paraImporter = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedFieldParagraph = paraImporter.ImportNode(fieldParagraph, true);
        resultBody.AppendChild(importedFieldParagraph);

        // ------------------------------------------------------------
        // 6. Save the extracted content.
        // ------------------------------------------------------------
        const string resultPath = "extracted-range.docx";
        resultDoc.Save(resultPath);

        // ------------------------------------------------------------
        // 7. Validate that the output file was created.
        // ------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Extraction output file was not created.");
    }
}
