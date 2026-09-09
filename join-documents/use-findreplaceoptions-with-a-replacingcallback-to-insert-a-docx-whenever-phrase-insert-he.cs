using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare file names.
        string mainDocPath = "MainDocument.docx";
        string insertDocPath = "InsertDocument.docx";
        string outputPath = "ResultDocument.docx";

        // Create the main document containing the placeholder phrase.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the beginning of the main document.");
        mainBuilder.Writeln("INSERT_HERE"); // Placeholder to be replaced.
        mainBuilder.Writeln("This is the end of the main document.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // Create the document that will be inserted.
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Document Start ===");
        insertBuilder.Writeln("This content comes from the inserted document.");
        insertBuilder.Writeln("=== Inserted Document End ===");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // Load the main document for processing.
        Document src = new Document(mainDocPath);

        // Set up FindReplaceOptions with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertDocumentAtReplaceHandler(insertDocPath);

        // Perform the replace operation using a regular expression that matches the placeholder.
        src.Range.Replace(new Regex("INSERT_HERE"), "", options);

        // Save the resulting document.
        src.Save(outputPath, SaveFormat.Docx);

        // Simple validation that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The result document was not saved correctly.");
    }

    // Callback that inserts a document at the location of each match.
    private class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        private readonly string _docToInsertPath;

        public InsertDocumentAtReplaceHandler(string docToInsertPath)
        {
            _docToInsertPath = docToInsertPath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document subDoc = new Document(_docToInsertPath);

            // The match is inside a Run; its parent paragraph is the insertion point.
            Paragraph para = (Paragraph)args.MatchNode.ParentNode;

            // Insert the document after the paragraph containing the placeholder.
            InsertDocument(para, subDoc);

            // Remove the placeholder paragraph.
            para.Remove();

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of docToInsert after the specified insertion destination.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node must be a paragraph or a table.");

            CompositeNode dstStory = insertionDestination.ParentNode;

            NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

            foreach (Section srcSection in docToInsert.Sections)
            {
                foreach (Node srcNode in srcSection.Body)
                {
                    // Skip the last empty paragraph of a section.
                    if (srcNode.NodeType == NodeType.Paragraph)
                    {
                        Paragraph p = (Paragraph)srcNode;
                        if (p.IsEndOfSection && !p.HasChildNodes)
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
