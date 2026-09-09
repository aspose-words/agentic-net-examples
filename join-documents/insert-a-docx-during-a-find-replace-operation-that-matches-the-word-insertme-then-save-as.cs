using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare file paths.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "JoinDocsSample");
        Directory.CreateDirectory(baseDir);
        string mainDocPath = Path.Combine(baseDir, "MainDocument.docx");
        string insertDocPath = Path.Combine(baseDir, "InsertDocument.docx");
        string outputPath = Path.Combine(baseDir, "ResultDocument.docx");

        // Create the main document containing the placeholder word "INSERTME".
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the beginning of the main document.");
        mainBuilder.Writeln("INSERTME"); // Placeholder to be replaced.
        mainBuilder.Writeln("This is the end of the main document.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // Create the document that will be inserted.
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Document Start ===");
        insertBuilder.Writeln("Hello from the inserted document!");
        insertBuilder.Writeln("=== Inserted Document End ===");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // Load the main document for processing.
        Document processingDoc = new Document(mainDocPath);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentAtReplaceHandler(insertDocPath)
        };

        // Perform the replace operation. The matched word will be removed and the new document inserted.
        processingDoc.Range.Replace(new Regex(@"\bINSERTME\b"), string.Empty, options);

        // Save the resulting document.
        processingDoc.Save(outputPath, SaveFormat.Docx);
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

            // The match is inside a paragraph; insert after that paragraph.
            Paragraph para = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(para, subDoc);

            // Remove the placeholder paragraph.
            para.Remove();

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of another document after a paragraph or table.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node must be a paragraph or table.");

            CompositeNode dstStory = insertionDestination.ParentNode;

            NodeImporter importer = new NodeImporter(
                docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

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
