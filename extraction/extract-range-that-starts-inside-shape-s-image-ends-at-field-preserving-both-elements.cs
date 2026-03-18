using System;
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

        // Insert an inline image shape.
        // The shape itself will be the start of the range we want to extract.
        Shape shape = builder.InsertImage(@"C:\Images\Sample.png");

        // Insert a field after the image. This field will be the end of the range.
        // Use any field type; here we use a simple MERGEFIELD.
        Field field = builder.InsertField("MERGEFIELD MyField");

        // Identify the start node (the shape) and the end node (the field end node).
        Node startNode = shape;
        Node endNode = field.End;

        // Create a new empty document that will hold the extracted range.
        Document extractedDoc = new Document();

        // Use a NodeImporter to copy nodes from the source document to the new document.
        NodeImporter importer = new NodeImporter(doc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        // Walk the node tree from the start node to the end node, importing each node.
        Node current = startNode;
        while (current != null)
        {
            // Import the current node (deep clone) into the target document.
            Node importedNode = importer.ImportNode(current, true);
            // Append the imported node to the body of the new document.
            extractedDoc.FirstSection.Body.AppendChild(importedNode);

            // Stop when we have imported the end node.
            if (current == endNode)
                break;

            // Move to the next sibling in the source document.
            current = current.NextSibling;
        }

        // Save the extracted document, preserving both the image shape and the field.
        extractedDoc.Save(@"C:\Output\ExtractedRange.docx");
    }
}
