using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace DocumentInsertionExample
{
    // Custom visitor that replaces a placeholder run with the contents of another document.
    class InsertDocVisitor : DocumentVisitor
    {
        private readonly Document _sourceDocument;
        private readonly NodeImporter _importer;

        public InsertDocVisitor(Document sourceDocument, Document destinationDocument)
        {
            _sourceDocument = sourceDocument;
            // Prepare an importer to handle style and list translation between the two documents.
            _importer = new NodeImporter(_sourceDocument, destinationDocument, ImportFormatMode.KeepSourceFormatting);
        }

        // Look for a run that contains the exact placeholder text "[InsertHere]".
        public override VisitorAction VisitRun(Run run)
        {
            if (run.Text == "[InsertHere]")
            {
                // The run's parent is a Paragraph. We'll insert after this paragraph.
                Paragraph placeholderParagraph = (Paragraph)run.ParentNode;
                Node insertionPoint = placeholderParagraph;

                // Import each block‑level node from the source document and insert it after the placeholder.
                foreach (Section srcSection in _sourceDocument.Sections)
                {
                    foreach (Node srcNode in srcSection.Body)
                    {
                        // Skip the final empty paragraph that Word adds to each section.
                        if (srcNode.NodeType == NodeType.Paragraph)
                        {
                            Paragraph para = (Paragraph)srcNode;
                            if (para.IsEndOfSection && !para.HasChildNodes)
                                continue;
                        }

                        Node importedNode = _importer.ImportNode(srcNode, true);
                        insertionPoint.ParentNode.InsertAfter(importedNode, insertionPoint);
                        insertionPoint = importedNode;
                    }
                }

                // Remove the placeholder run (and its paragraph if it becomes empty).
                run.Remove();
                if (!placeholderParagraph.HasChildNodes)
                    placeholderParagraph.Remove();

                // No need to visit children of this run any further.
                // Run nodes have no children, so returning Continue is sufficient.
                return VisitorAction.Continue;
            }

            return VisitorAction.Continue;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the destination document that contains the placeholder.
            Document destDoc = new Document("Destination.docx");

            // Load the source document whose content we want to insert.
            Document srcDoc = new Document("Source.docx");

            // Create and run the visitor to perform the insertion.
            InsertDocVisitor visitor = new InsertDocVisitor(srcDoc, destDoc);
            destDoc.Accept(visitor);

            // Save the modified document.
            destDoc.Save("Result.docx");
        }
    }
}
