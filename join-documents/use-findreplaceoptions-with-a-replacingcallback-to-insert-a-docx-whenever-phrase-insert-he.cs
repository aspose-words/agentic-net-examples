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
        string baseDir = Directory.GetCurrentDirectory();
        string mainDocPath = Path.Combine(baseDir, "MainDocument.docx");
        string subDocPath = Path.Combine(baseDir, "InsertDocument.docx");
        string resultPath = Path.Combine(baseDir, "ResultDocument.docx");

        // Create the main document containing the placeholder.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the beginning of the main document.");
        mainBuilder.Writeln("INSERT_HERE"); // Placeholder to be replaced.
        mainBuilder.Writeln("This is the end of the main document.");
        mainDoc.Save(mainDocPath);

        // Create the document that will be inserted.
        Document subDoc = new Document();
        DocumentBuilder subBuilder = new DocumentBuilder(subDoc);
        subBuilder.Writeln("This is the first line of the inserted document.");
        subBuilder.Writeln("This is the second line of the inserted document.");
        subDoc.Save(subDocPath);

        // Load the main document for processing.
        Document srcDoc = new Document(mainDocPath);

        // Configure find‑replace with a callback that inserts the sub‑document.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentAtReplaceHandler(subDocPath)
        };

        // Perform the replace operation. The placeholder text is removed by the callback.
        srcDoc.Range.Replace(new Regex("INSERT_HERE"), string.Empty, options);

        // Save the merged result.
        srcDoc.Save(resultPath, SaveFormat.Docx);

        // Verify that the result file was created.
        if (!File.Exists(resultPath))
        {
            throw new InvalidOperationException($"The result file was not created: {resultPath}");
        }
    }

    // Callback that inserts a document at the location of each match.
    private class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        private readonly string _documentPath;

        public InsertDocumentAtReplaceHandler(string documentPath)
        {
            _documentPath = documentPath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document insertDoc = new Document(_documentPath);

            // The match is inside a paragraph; insert after that paragraph.
            Paragraph paragraph = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(paragraph, insertDoc);

            // Remove the placeholder paragraph.
            paragraph.Remove();

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of another document after a paragraph or table.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node must be either a paragraph or a table.");

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
