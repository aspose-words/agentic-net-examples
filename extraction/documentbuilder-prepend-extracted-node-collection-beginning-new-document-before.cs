using System;
using Aspose.Words;

class PrependNodesExample
{
    static void Main()
    {
        // Create a source document with some sample content.
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("First paragraph in source document.");
        srcBuilder.Writeln("Second paragraph in source document.");
        srcDoc.EnsureMinimum();

        // Create a new empty document.
        Document newDoc = new Document();
        newDoc.EnsureMinimum();

        // Prepare a NodeImporter to import nodes from srcDoc to newDoc.
        NodeImporter importer = new NodeImporter(srcDoc, newDoc, ImportFormatMode.KeepSourceFormatting);

        // Get only paragraph nodes from the source document's body.
        NodeCollection srcParagraphs = srcDoc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true);

        // Insert the imported paragraphs at the beginning of the new document's body.
        Body newBody = newDoc.FirstSection.Body;
        for (int i = srcParagraphs.Count - 1; i >= 0; i--)
        {
            Node importedNode = importer.ImportNode(srcParagraphs[i], true);
            newBody.PrependChild(importedNode);
        }

        // Save the resulting document.
        newDoc.Save("Result.docx");
        Console.WriteLine("Result.docx created successfully.");
    }
}
