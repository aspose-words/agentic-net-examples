using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;

class ExtractShapeToFieldRange
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape instead of an external image.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.FillColor = System.Drawing.Color.LightBlue;

        // Insert a field after the shape. This field will be the end of the range.
        Field field = builder.InsertField("MERGEFIELD MyField");

        // The shape and the field are both inside the same paragraph.
        Paragraph paragraph = (Paragraph)shape.ParentNode;

        // Create a new empty document that will hold the extracted range.
        Document extractedDoc = new Document();

        // Use a NodeImporter to copy the paragraph from the source document to the new document.
        NodeImporter importer = new NodeImporter(doc, extractedDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(paragraph, true);
        extractedDoc.FirstSection.Body.AppendChild(importedParagraph);

        // Save the extracted document to a path relative to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ExtractedRange.docx");
        extractedDoc.Save(outputPath);

        Console.WriteLine($"Extracted document saved to: {outputPath}");
    }
}
