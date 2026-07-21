using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string dataDir = Path.Combine(Environment.CurrentDirectory, "Data");
        Directory.CreateDirectory(dataDir);

        // Create the main document that contains the placeholder.
        string mainDocPath = Path.Combine(dataDir, "Main.docx");
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the beginning of the document.");
        mainBuilder.Writeln("PLACEHOLDER"); // The phrase to be replaced.
        mainBuilder.Writeln("This is the end of the document.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // Create the document that will replace the placeholder.
        string insertDocPath = Path.Combine(dataDir, "Insert.docx");
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is the inserted document.");
        insertBuilder.Writeln("It has multiple paragraphs.");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // Load the main document for processing.
        Document srcDoc = new Document(mainDocPath);

        // Set up find‑replace with a callback that inserts a document.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertDocumentAtReplaceHandler(insertDocPath);

        // Replace the placeholder text. The replacement string is empty because the callback handles insertion.
        srcDoc.Range.Replace(new Regex("\\bPLACEHOLDER\\b"), string.Empty, options);

        // Save the resulting document as ODT.
        string resultPath = Path.Combine(dataDir, "Result.odt");
        srcDoc.Save(resultPath, SaveFormat.Odt);

        // Validate that the ODT file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The ODT file was not created.");
    }

    // Callback that inserts a document at the location of each match.
    private class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        private readonly string _docPath;

        public InsertDocumentAtReplaceHandler(string docPath)
        {
            _docPath = docPath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document subDoc = new Document(_docPath);

            // The match is inside a paragraph; insert the sub‑document after this paragraph.
            Paragraph para = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(para, subDoc);

            // Remove the placeholder paragraph.
            para.Remove();

            // Skip the default replace action because we have already handled the insertion.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of subDoc after the specified paragraph or table.
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
