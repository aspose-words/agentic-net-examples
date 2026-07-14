using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Prepare a temporary working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // File names for the source documents and the final merged document.
        string mainDocPath = Path.Combine(workDir, "MainDocument.docx");
        string insertDocPath = Path.Combine(workDir, "InsertDocument.docx");
        string resultDocPath = Path.Combine(workDir, "ResultDocument.docx");

        // -----------------------------------------------------------------
        // 1. Create the main document that contains the placeholder text.
        // -----------------------------------------------------------------
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the beginning of the main document.");
        mainBuilder.Writeln("INSERT_HERE"); // Placeholder to be replaced.
        mainBuilder.Writeln("This is the end of the main document.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // 2. Create the document whose content will be inserted.
        // ---------------------------------------------------------------
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Document Start ===");
        insertBuilder.Writeln("Here is some content that will be inserted.");
        insertBuilder.Writeln("=== Inserted Document End ===");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // 3. Load the main document and perform find‑replace with a callback.
        // ---------------------------------------------------------------
        Document loadedMain = new Document(mainDocPath);
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentAtReplaceHandler(insertDocPath)
        };

        // The placeholder is replaced by an empty string; the callback inserts the document.
        loadedMain.Range.Replace(new Regex("INSERT_HERE"), string.Empty, options);

        // ---------------------------------------------------------------
        // 4. Save the resulting document.
        // ---------------------------------------------------------------
        loadedMain.Save(resultDocPath, SaveFormat.Docx);

        // Simple validation to ensure the file was created.
        if (!File.Exists(resultDocPath))
            throw new InvalidOperationException("The result document was not saved correctly.");

        // The program finishes without waiting for user input.
    }

    // -----------------------------------------------------------------
    // Callback class that inserts a document at the location of each match.
    // -----------------------------------------------------------------
    private class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        private readonly string _documentToInsertPath;

        public InsertDocumentAtReplaceHandler(string documentToInsertPath)
        {
            _documentToInsertPath = documentToInsertPath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document that will be inserted.
            Document subDoc = new Document(_documentToInsertPath);

            // The match is inside a paragraph; insert after that paragraph.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(placeholderParagraph, subDoc);

            // Remove the paragraph that contained the placeholder text.
            placeholderParagraph.Remove();

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of the supplied document after the specified paragraph.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("Insertion destination must be a paragraph or a table.");

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
