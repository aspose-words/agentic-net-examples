using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace DocumentJoinExample
{
    // Custom visitor that copies the content of a source document into a target document
    // using a DocumentBuilder and a NodeImporter.
    class InsertDocumentVisitor : DocumentVisitor
    {
        private readonly DocumentBuilder _builder;
        private readonly NodeImporter _importer;

        public InsertDocumentVisitor(DocumentBuilder builder, Document source)
        {
            _builder = builder;
            // Importer handles style and formatting conflicts.
            _importer = new NodeImporter(source, builder.Document, ImportFormatMode.KeepSourceFormatting);
        }

        // Paragraphs: import the whole paragraph node.
        public override VisitorAction VisitParagraphStart(Paragraph paragraph)
        {
            Node imported = _importer.ImportNode(paragraph, true);
            _builder.InsertNode(imported);
            // Move the cursor into the imported paragraph so that any nested nodes (e.g., runs) are added correctly.
            _builder.MoveTo(imported);
            return VisitorAction.Continue;
        }

        // Tables: import the whole table node.
        public override VisitorAction VisitTableStart(Table table)
        {
            Node imported = _importer.ImportNode(table, true);
            _builder.InsertNode(imported);
            _builder.MoveTo(imported);
            return VisitorAction.Continue;
        }

        // Runs: write the text directly (runs are already part of the imported paragraph,
        // but handling them ensures that isolated runs are also copied).
        public override VisitorAction VisitRun(Run run)
        {
            _builder.Write(run.Text);
            return VisitorAction.Continue;
        }

        // For any other node types that are not explicitly handled, import them as is.
        // DocumentVisitor does not provide a generic "VisitNodeStart" method, so we handle
        // the remaining node types by overriding the specific start methods that are
        // available in the API (e.g., VisitHeaderFooterStart, VisitShapeStart, etc.).
        // In this example we simply ignore them because the most common content types
        // (paragraphs, tables, runs) are already covered.
    }

    class Program
    {
        static void Main()
        {
            // Create the destination (combined) document.
            Document dstDoc = new Document(); // blank document
            DocumentBuilder builder = new DocumentBuilder(dstDoc);
            builder.Writeln("=== Start of Combined Document ===");

            // Load the source document that we want to insert.
            Document srcDoc = new Document("Source.docx"); // replace with actual path

            // Use the custom visitor to copy the source content into the destination.
            InsertDocumentVisitor visitor = new InsertDocumentVisitor(builder, srcDoc);
            srcDoc.Accept(visitor);

            builder.Writeln("=== End of Combined Document ===");

            // Save the combined document.
            dstDoc.Save("Combined.docx"); // replace with desired output path
        }
    }
}
