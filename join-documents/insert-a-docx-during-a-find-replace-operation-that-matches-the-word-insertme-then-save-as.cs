using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary documents.
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsFolder);

        string mainPath = Path.Combine(docsFolder, "Main.docx");
        string insertPath = Path.Combine(docsFolder, "Insert.docx");
        string resultPath = Path.Combine(docsFolder, "Result.docx");

        // Create the main document containing the placeholder word "INSERTME".
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the main document.");
        mainBuilder.Writeln("INSERTME"); // placeholder to be replaced.
        mainBuilder.Writeln("End of the main document.");
        mainDoc.Save(mainPath, SaveFormat.Docx);

        // Create the document that will be inserted.
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is the inserted document content.");
        insertDoc.Save(insertPath, SaveFormat.Docx);

        // Load the main document for processing.
        Document srcDoc = new Document(mainPath);

        // Configure find‑replace with a callback that inserts the document.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertDocumentHandler(insertPath);

        // Perform the replace operation; the placeholder will be replaced by the inserted document.
        srcDoc.Range.Replace(new Regex("\\bINSERTME\\b"), "", options);

        // Save the resulting document.
        srcDoc.Save(resultPath, SaveFormat.Docx);

        // Simple validation to ensure the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The result document was not created.");
    }

    // Callback that inserts a document at the location of the matched text.
    private class InsertDocumentHandler : IReplacingCallback
    {
        private readonly string _docPath;

        public InsertDocumentHandler(string docPath)
        {
            _docPath = docPath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document subDoc = new Document(_docPath);

            // The match node is a Run; its parent is the Paragraph containing the placeholder.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;

            // Insert the sub‑document after the placeholder paragraph.
            InsertDocument(placeholderParagraph, subDoc);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip the default replacement (we already handled it).
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of docToInsert after insertionDestination (paragraph or table).
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node must be a paragraph or table.");

            CompositeNode dstStory = insertionDestination.ParentNode;
            NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

            foreach (Section srcSection in docToInsert.Sections)
            {
                foreach (Node srcNode in srcSection.Body)
                {
                    // Skip the final empty paragraph of a section.
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
