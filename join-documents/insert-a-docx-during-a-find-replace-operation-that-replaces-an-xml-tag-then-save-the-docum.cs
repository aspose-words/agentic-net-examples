using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Paths for the documents.
        string mainDocPath = Path.Combine(dataDir, "Main.docx");
        string insertDocPath = Path.Combine(dataDir, "Insert.docx");
        string resultPath = Path.Combine(dataDir, "Result.docx");

        // -----------------------------------------------------------------
        // Create the main document that contains the XML tag to be replaced.
        // -----------------------------------------------------------------
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("Document start.");
        // The placeholder tag that will be searched for.
        mainBuilder.Writeln("<myTag/>");
        mainBuilder.Writeln("Document end.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // Create the document whose content will be inserted during replace.
        // ---------------------------------------------------------------
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Content Start ===");
        insertBuilder.Writeln("Hello from the inserted document.");
        insertBuilder.Writeln("=== Inserted Content End ===");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // Load the main document and perform a find‑replace with a callback.
        // ---------------------------------------------------------------
        Document srcDoc = new Document(mainDocPath);
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentCallback(insertDocPath)
        };

        // Use a regular expression to locate the exact XML tag.
        srcDoc.Range.Replace(new Regex("<myTag/>"), string.Empty, options);
        srcDoc.Save(resultPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // Verify that the result file was created.
        // ---------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The result document was not created.");
    }

    // Callback that inserts a document at the location of the matched tag.
    private class InsertDocumentCallback : IReplacingCallback
    {
        private readonly string _docPath;

        public InsertDocumentCallback(string docPath)
        {
            _docPath = docPath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document subDoc = new Document(_docPath);

            // The match resides inside a paragraph.
            Paragraph placeholderParagraph = args.MatchNode.ParentNode as Paragraph;
            if (placeholderParagraph != null)
            {
                InsertDocument(placeholderParagraph, subDoc);
                // Remove the paragraph that contained the placeholder tag.
                placeholderParagraph.Remove();
            }

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of docToInsert after insertionDestination.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph &&
                insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("Insertion destination must be a paragraph or table.");

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
