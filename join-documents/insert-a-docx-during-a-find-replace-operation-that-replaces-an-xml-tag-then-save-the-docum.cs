using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for temporary documents.
        string mainDocPath = Path.Combine(outputDir, "Main.docx");
        string insertDocPath = Path.Combine(outputDir, "Insert.docx");
        string resultDocPath = Path.Combine(outputDir, "Result.docx");

        // Create the main document containing a placeholder tag.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("Document start.");
        mainBuilder.Writeln("[INSERT_DOC]"); // Placeholder to be replaced.
        mainBuilder.Writeln("Document end.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // Create the document that will be inserted.
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is the inserted document.");
        insertBuilder.Writeln("It has multiple paragraphs.");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // Load the main document for processing.
        Document srcDoc = new Document(mainDocPath);

        // Set up find‑replace with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertDocumentHandler(insertDocPath);

        // Perform the replace; the placeholder will be removed and the insertDoc inserted.
        srcDoc.Range.Replace(new Regex(@"\[INSERT_DOC\]"), string.Empty, options);

        // Save the final merged document.
        srcDoc.Save(resultDocPath, SaveFormat.Docx);

        // Simple validation to ensure the file was created.
        if (!File.Exists(resultDocPath))
            throw new InvalidOperationException("Result document was not created.");
    }

    // Callback that inserts a document at the location of the matched text.
    private class InsertDocumentHandler : IReplacingCallback
    {
        private readonly string _docToInsertPath;

        public InsertDocumentHandler(string docToInsertPath)
        {
            _docToInsertPath = docToInsertPath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document docToInsert = new Document(_docToInsertPath);

            // The match is inside a paragraph; get that paragraph.
            Paragraph placeholderParagraph = args.MatchNode.ParentNode as Paragraph;
            if (placeholderParagraph == null)
                return ReplaceAction.Skip;

            // Insert the document after the placeholder paragraph.
            InsertDocument(placeholderParagraph, docToInsert);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of docToInsert after the specified paragraph.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("Insertion destination must be a paragraph or table.");

            CompositeNode dstStory = insertionDestination.ParentNode;
            NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

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

                    Node newNode = importer.ImportNode(srcNode, true);
                    dstStory.InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
    }
}
