using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // ---------- Create a sample document ----------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph that will contain the image.
        builder.Writeln();

        // Insert a tiny PNG image (1x1 pixel) from a Base64 string.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imgStream);
        }

        // End the paragraph that holds the image.
        builder.Writeln();

        // Insert a DATE field in the next paragraph.
        builder.InsertField(@"DATE \@ ""MMMM d, yyyy""");
        builder.Writeln();

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // ---------- Load the document ----------
        Document loadedDoc = new Document(sourcePath);

        // Locate the image‑bearing shape.
        Shape imageShape = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                    .OfType<Shape>()
                                    .FirstOrDefault(s => s.HasImage);
        if (imageShape == null)
            throw new InvalidOperationException("Image shape not found.");

        // Locate the first field (the DATE field we inserted).
        Field field = loadedDoc.Range.Fields.FirstOrDefault();
        if (field == null)
            throw new InvalidOperationException("Field not found.");

        // Determine the paragraphs that contain the shape and the field.
        Paragraph shapeParagraph = imageShape.GetAncestor(NodeType.Paragraph) as Paragraph;
        Paragraph fieldParagraph = field.Start.GetAncestor(NodeType.Paragraph) as Paragraph;
        if (shapeParagraph == null || fieldParagraph == null)
            throw new InvalidOperationException("Unable to locate containing paragraphs.");

        // ---------- Extract the range (shape paragraph through field paragraph) ----------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        // Create a new section and body for the result document.
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to import nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        // Import and add the shape paragraph.
        Node importedShapePara = importer.ImportNode(shapeParagraph, true);
        resultBody.AppendChild(importedShapePara);

        // If the field resides in a different paragraph, import and add it as well.
        if (!ReferenceEquals(shapeParagraph, fieldParagraph))
        {
            Node importedFieldPara = importer.ImportNode(fieldParagraph, true);
            resultBody.AppendChild(importedFieldPara);
        }

        // Save the extracted content.
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // ---------- Validation ----------
        // Verify that the extracted document still contains an image shape.
        Shape extractedShape = resultDoc.GetChildNodes(NodeType.Shape, true)
                                        .OfType<Shape>()
                                        .FirstOrDefault(s => s.HasImage);
        if (extractedShape == null)
            throw new InvalidOperationException("Extracted image shape is missing.");

        // Verify that the extracted document still contains a field.
        if (!resultDoc.Range.Fields.Any())
            throw new InvalidOperationException("Extracted field is missing.");

        Console.WriteLine("Extraction completed successfully.");
    }
}
