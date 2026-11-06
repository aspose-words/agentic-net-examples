using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class ExtractionExample
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document containing an image (inside a shape) and a field.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before image.");

        // Insert a tiny PNG image (inline shape).
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imgStream);
        }

        builder.Writeln("Paragraph after image, before field.");

        // Insert a DATE field.
        builder.InsertField(FieldType.FieldDate, true);
        builder.Writeln("Paragraph after field.");

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document and locate the start (image) and end (field) nodes.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        Shape imageShape = loadedDoc.GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .FirstOrDefault(s => s.HasImage);
        if (imageShape == null)
            throw new InvalidOperationException("No image shape found in the document.");

        FieldStart fieldStart = loadedDoc.GetChildNodes(NodeType.FieldStart, true)
            .OfType<FieldStart>()
            .FirstOrDefault();
        if (fieldStart == null)
            throw new InvalidOperationException("No field start node found in the document.");

        Paragraph startParagraph = imageShape.ParentNode as Paragraph;
        Paragraph endParagraph = fieldStart.ParentNode as Paragraph;
        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Unable to locate parent paragraphs for extraction.");

        // -----------------------------------------------------------------
        // 3. Build a new document that will contain the extracted range.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to copy nodes from the source document to the result document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        Node importedStart = importer.ImportNode(startParagraph, true);
        resultBody.AppendChild(importedStart);

        if (!ReferenceEquals(startParagraph, endParagraph))
        {
            Node importedEnd = importer.ImportNode(endParagraph, true);
            resultBody.AppendChild(importedEnd);
        }

        // -----------------------------------------------------------------
        // 4. Save the extracted content.
        // -----------------------------------------------------------------
        const string resultPath = "extracted-range.docx";
        resultDoc.Save(resultPath);

        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");

        // -----------------------------------------------------------------
        // 5. Optional: write a JSON report about the extraction.
        // -----------------------------------------------------------------
        var report = new
        {
            SourceDocument = sourcePath,
            ExtractedDocument = resultPath,
            ExtractedElements = new[] { "ImageShape", "Field" },
            ExtractionTimeUtc = DateTime.UtcNow
        };

        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("extraction-report.json", jsonReport);
    }
}
