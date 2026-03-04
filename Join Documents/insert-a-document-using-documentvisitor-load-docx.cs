using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace DocumentInsertionExample
{
    // Custom visitor that copies block‑level nodes from the source document
    // into a destination document at a specified insertion point.
    class InsertVisitor : DocumentVisitor
    {
        private readonly Document _destDocument;
        private Node _insertionNode;               // Node after which new content will be inserted.
        private readonly NodeImporter _importer;    // Re‑uses import logic for efficiency.
        private bool _skipChildren;                // Used when the current node has been imported as a whole.

        public InsertVisitor(Document destDocument, Node insertionNode, Document sourceDocument)
        {
            _destDocument = destDocument;
            _insertionNode = insertionNode;
            // Keep source formatting when importing.
            _importer = new NodeImporter(sourceDocument, destDocument, ImportFormatMode.KeepSourceFormatting);
            _skipChildren = false;
        }

        // Copy paragraphs.
        public override VisitorAction VisitParagraphStart(Paragraph paragraph)
        {
            if (_skipChildren)
                return VisitorAction.Continue; // Paragraph is part of a table/shape that has already been imported.

            Node importedNode = _importer.ImportNode(paragraph, true);
            _insertionNode.ParentNode.InsertAfter(importedNode, _insertionNode);
            _insertionNode = importedNode;
            return VisitorAction.Continue;
        }

        // Copy tables.
        public override VisitorAction VisitTableStart(Table table)
        {
            // Import the whole table (including its rows, cells, paragraphs, etc.).
            Node importedNode = _importer.ImportNode(table, true);
            _insertionNode.ParentNode.InsertAfter(importedNode, _insertionNode);
            _insertionNode = importedNode;

            // The table's children have already been copied, so tell the visitor to ignore them.
            _skipChildren = true;
            return VisitorAction.Continue;
        }

        public override VisitorAction VisitTableEnd(Table table)
        {
            // Reset the flag so that following nodes are processed normally.
            _skipChildren = false;
            return VisitorAction.Continue;
        }

        // Copy shapes (e.g., images, drawings).
        public override VisitorAction VisitShapeStart(Shape shape)
        {
            Node importedNode = _importer.ImportNode(shape, true);
            _insertionNode.ParentNode.InsertAfter(importedNode, _insertionNode);
            _insertionNode = importedNode;

            // Shapes may contain child nodes (e.g., paragraphs inside a text box). Skip them.
            _skipChildren = true;
            return VisitorAction.Continue;
        }

        public override VisitorAction VisitShapeEnd(Shape shape)
        {
            _skipChildren = false;
            return VisitorAction.Continue;
        }

        // For all other node types we simply continue without special handling.
        public override VisitorAction VisitRun(Run run) => VisitorAction.Continue;
        public override VisitorAction VisitSectionStart(Section section) => VisitorAction.Continue;
        public override VisitorAction VisitBodyStart(Body body) => VisitorAction.Continue;
        public override VisitorAction VisitDocumentStart(Document doc) => VisitorAction.Continue;
    }

    class Program
    {
        static void Main()
        {
            // Load the destination document (the one we will insert into).
            Document destDoc = new Document("Destination.docx");

            // Load the source document (the content to be inserted).
            Document srcDoc = new Document("Source.docx");

            // Prepare the insertion point – for this example we use a bookmark named "InsertHere".
            DocumentBuilder builder = new DocumentBuilder(destDoc);
            builder.MoveToBookmark("InsertHere");

            // The current paragraph is the node after which new content will be placed.
            Node insertionPoint = builder.CurrentParagraph;

            // Create and run the visitor on the source document.
            InsertVisitor visitor = new InsertVisitor(destDoc, insertionPoint, srcDoc);
            srcDoc.Accept(visitor);

            // Save the modified destination document.
            destDoc.Save("Result.docx");
        }
    }
}
