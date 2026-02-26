using System;
using Aspose.Words;

class DocumentJoiner
{
    static void Main()
    {
        // Paths to the source and destination DOCX files.
        string sourcePath = "Source.docx";
        string destinationPath = "Destination.docx";
        string outputPath = "Joined.docx";

        // Load the destination document (the document into which we will insert the source).
        Document dstDoc = new Document(destinationPath);

        // Load the source document (the document to be inserted).
        Document srcDoc = new Document(sourcePath);

        // Determine the node after which the imported nodes will be inserted.
        // If the destination document is empty we start after the first paragraph.
        Node insertionDestination = dstDoc.LastSection?.Body?.LastChild ??
                                    dstDoc.FirstSection?.Body?.FirstParagraph;

        // Initialise the NodeImporter with ImportFormatMode.UseDestinationStyles.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.UseDestinationStyles);

        // Import each block‑level node from the source document and insert it after the current destination node.
        foreach (Section srcSection in srcDoc.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Skip the automatically added empty paragraph at the end of a section.
                if (srcNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)srcNode;
                    if (para.IsEndOfSection && !para.HasChildNodes)
                        continue;
                }

                // Import the node (deep clone) into the destination document.
                Node importedNode = importer.ImportNode(srcNode, true);

                // Insert the imported node after the previously inserted node.
                insertionDestination.ParentNode.InsertAfter(importedNode, insertionDestination);
                insertionDestination = importedNode;
            }
        }

        // Save the combined document.
        dstDoc.Save(outputPath);
    }
}
