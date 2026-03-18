using System;
using Aspose.Words;

class PrependNodesExample
{
    static void Main()
    {
        // Load the source document whose nodes we want to extract.
        Document srcDoc = new Document("Source.docx");

        // Create a new empty document and attach a DocumentBuilder to it.
        Document newDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(newDoc);

        // Ensure the new document has at least one section/body.
        newDoc.EnsureMinimum();

        // Move the builder's cursor to the very start of the document.
        builder.MoveToDocumentStart();

        // Prepare a NodeImporter to efficiently import nodes from srcDoc to newDoc.
        NodeImporter importer = new NodeImporter(srcDoc, newDoc, ImportFormatMode.KeepSourceFormatting);

        // Get all child nodes (paragraphs, tables, etc.) from the first section's body of the source.
        NodeCollection srcNodes = srcDoc.FirstSection.Body.GetChildNodes(NodeType.Any, true);

        // Insert the imported nodes at the beginning of the new document.
        // InsertNode inserts before the current cursor position, so we iterate in reverse order
        // to preserve the original ordering of the source nodes.
        for (int i = srcNodes.Count - 1; i >= 0; i--)
        {
            Node importedNode = importer.ImportNode(srcNodes[i], true);
            builder.InsertNode(importedNode);
        }

        // Save the resulting document.
        newDoc.Save("Result.docx");
    }
}
