using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the main document that contains a content control (StructuredDocumentTag)
        Document mainDoc = new Document(@"C:\Docs\MainDocument.docx");

        // Load the document that we want to insert into the content control
        Document subDoc = new Document(@"C:\Docs\SubDocument.docx");

        // Find the content control by its title (or any other property you prefer)
        StructuredDocumentTag targetSdt = mainDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .Cast<StructuredDocumentTag>()
            .FirstOrDefault(tag => tag.Title == "InsertHere");

        // Ensure the content control was found
        if (targetSdt == null)
        {
            Console.WriteLine("Content control with the specified title was not found.");
            return;
        }

        // Use NodeImporter for efficient import of nodes from the sub‑document
        NodeImporter importer = new NodeImporter(subDoc, mainDoc, ImportFormatMode.KeepSourceFormatting);

        // Import each node from the sub‑document into the content control.
        // We can append directly to the StructuredDocumentTag – it will place the nodes inside its SdtContent.
        foreach (Section srcSection in subDoc.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Skip the final empty paragraph that Word adds to each section
                if (srcNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)srcNode;
                    if (para.IsEndOfSection && !para.HasChildNodes)
                        continue;
                }

                // Import the node into the destination document
                Node importedNode = importer.ImportNode(srcNode, true);

                // Append the imported node to the content control
                targetSdt.AppendChild(importedNode);
            }
        }

        // Save the updated document
        mainDoc.Save(@"C:\Docs\ResultDocument.docx");
    }
}
