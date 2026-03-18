using System;
using Aspose.Words;
using Aspose.Words.Layout;

class Program
{
    static void Main()
    {
        // Load the source document that will be inserted.
        Document srcDoc = new Document("Source.docx");

        // Load the destination document where the source will be inserted.
        Document dstDoc = new Document("Destination.docx");

        // Insert the source document after the last paragraph of the destination.
        InsertDocument(dstDoc.LastSection.Body.LastParagraph, srcDoc);

        // Validate that each inserted section has correct page numbers.
        ValidateInsertedSections(dstDoc);

        // Save the resulting document.
        dstDoc.Save("Result.docx");
    }

    // Inserts the contents of a document after the specified node.
    static void InsertDocument(Node insertionDestination, Document docToInsert)
    {
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("The destination node should be either a paragraph or table.");

        CompositeNode destinationParent = insertionDestination.ParentNode;

        // Use NodeImporter for efficient import of nodes.
        NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

        // Loop through all sections and their body nodes in the source document.
        foreach (Section srcSection in docToInsert.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Skip the last empty paragraph of a section.
                if (srcNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)srcNode;
                    if (para.IsEndOfSection && !para.HasChildNodes)
                        continue;
                }

                // Import the node into the destination document.
                Node newNode = importer.ImportNode(srcNode, true);

                // Insert the imported node after the current insertion point.
                destinationParent.InsertAfter(newNode, insertionDestination);
                insertionDestination = newNode;
            }
        }
    }

    // Validates that each section in the document has a sensible page range.
    static void ValidateInsertedSections(Document doc)
    {
        // Ensure the layout information is up to date.
        doc.UpdatePageLayout();

        // LayoutCollector maps nodes to page numbers.
        LayoutCollector layoutCollector = new LayoutCollector(doc);

        // Iterate through all sections in the document.
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            Section section = doc.Sections[i];

            // Get the first and last paragraphs of the section body.
            Paragraph firstParagraph = section.Body.FirstParagraph;
            Paragraph lastParagraph = section.Body.LastParagraph;

            // Retrieve the start and end page numbers for the section.
            int startPage = layoutCollector.GetStartPageIndex(firstParagraph);
            int endPage = layoutCollector.GetEndPageIndex(lastParagraph);

            // Simple validation: start page must be <= end page and both must be non‑zero.
            if (startPage == 0 || endPage == 0)
            {
                Console.WriteLine($"Section {i + 1}: Unable to determine page numbers.");
            }
            else if (startPage > endPage)
            {
                Console.WriteLine($"Section {i + 1}: Invalid page range (start {startPage} > end {endPage}).");
            }
            else
            {
                Console.WriteLine($"Section {i + 1}: Pages {startPage} to {endPage}.");
            }
        }
    }
}
