using System;
using Aspose.Words;
// Note: In recent Aspose.Words versions NodeImporter lives in Aspose.Words.Importing.
// If you are using an older version the class is directly under Aspose.Words, so we omit the Importing namespace.

class DocumentNodeImporterDemo
{
    static void Main()
    {
        // Load the source and destination documents from disk.
        Document srcDoc = new Document("Source.docx");
        Document dstDocUseDest = new Document("Destination.docx");
        Document dstDocKeepSource = new Document("Destination.docx");
        Document dstDocKeepDiff = new Document("Destination.docx");

        // Insert using ImportFormatMode.UseDestinationStyles.
        InsertDocumentUsingMode(srcDoc, dstDocUseDest, ImportFormatMode.UseDestinationStyles);
        dstDocUseDest.Save("Result_UseDestinationStyles.docx");

        // Insert using ImportFormatMode.KeepSourceFormatting.
        InsertDocumentUsingMode(srcDoc, dstDocKeepSource, ImportFormatMode.KeepSourceFormatting);
        dstDocKeepSource.Save("Result_KeepSourceFormatting.docx");

        // Insert using ImportFormatMode.KeepDifferentStyles.
        InsertDocumentUsingMode(srcDoc, dstDocKeepDiff, ImportFormatMode.KeepDifferentStyles);
        dstDocKeepDiff.Save("Result_KeepDifferentStyles.docx");
    }

    // Helper method that imports all block‑level nodes from srcDoc into dstDoc using the specified mode.
    static void InsertDocumentUsingMode(Document srcDoc, Document dstDoc, ImportFormatMode mode)
    {
        // Create a NodeImporter that will handle style and list translation.
        // The fully‑qualified name works for both old and new Aspose.Words versions.
        var importer = new NodeImporter(srcDoc, dstDoc, mode);

        // Determine the node after which we will insert the imported content.
        // Here we use the last node of the destination document's body.
        Node insertionPoint = dstDoc.FirstSection.Body.LastChild;

        // Iterate through each section and each node in the source document.
        foreach (Section srcSection in srcDoc.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Skip the final empty paragraph of a section (Word adds it automatically).
                if (srcNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)srcNode;
                    if (para.IsEndOfSection && !para.HasChildNodes)
                        continue;
                }

                // Import the node (deep clone) into the destination document.
                Node importedNode = importer.ImportNode(srcNode, true);

                // Insert the imported node after the current insertion point.
                dstDoc.FirstSection.Body.InsertAfter(importedNode, insertionPoint);
                insertionPoint = importedNode;
            }
        }
    }
}
