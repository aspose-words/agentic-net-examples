using System;
using Aspose.Words;
using Aspose.Words.Layout;

class Program
{
    static void Main()
    {
        // Create source document with sample content.
        Document srcDoc = CreateSampleDocument("Source", new[]
        {
            "First source paragraph.",
            "Second source paragraph."
        });

        // Create destination document with sample content.
        Document dstDoc = CreateSampleDocument("Destination", new[]
        {
            "First destination paragraph.",
            "Second destination paragraph."
        });

        // Insert the source document after the last paragraph of the destination.
        InsertDocument(dstDoc.LastSection.Body.LastParagraph, srcDoc);

        // Validate that each inserted section has correct page numbers.
        ValidateInsertedSections(dstDoc);

        // Save the resulting document (optional).
        dstDoc.Save("Result.docx");
    }

    static Document CreateSampleDocument(string title, string[] paragraphs)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln(title);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        foreach (var text in paragraphs)
        {
            builder.Writeln(text);
        }

        return doc;
    }

    // Inserts the contents of a document after the specified node.
    static void InsertDocument(Node insertionDestination, Document docToInsert)
    {
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("The destination node should be either a paragraph or table.");

        CompositeNode destinationParent = insertionDestination.ParentNode;

        NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

        foreach (Section srcSection in docToInsert.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                if (srcNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)srcNode;
                    if (para.IsEndOfSection && !para.HasChildNodes)
                        continue;
                }

                Node newNode = importer.ImportNode(srcNode, true);
                destinationParent.InsertAfter(newNode, insertionDestination);
                insertionDestination = newNode;
            }
        }
    }

    // Validates that each section in the document has a sensible page range.
    static void ValidateInsertedSections(Document doc)
    {
        doc.UpdatePageLayout();

        LayoutCollector layoutCollector = new LayoutCollector(doc);

        for (int i = 0; i < doc.Sections.Count; i++)
        {
            Section section = doc.Sections[i];
            Paragraph firstParagraph = section.Body.FirstParagraph;
            Paragraph lastParagraph = section.Body.LastParagraph;

            int startPage = layoutCollector.GetStartPageIndex(firstParagraph);
            int endPage = layoutCollector.GetEndPageIndex(lastParagraph);

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
