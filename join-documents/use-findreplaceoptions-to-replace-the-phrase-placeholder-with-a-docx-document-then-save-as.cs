using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the documents.
        string mainDocPath = Path.Combine(outputDir, "MainDocument.docx");
        string insertDocPath = Path.Combine(outputDir, "InsertDocument.docx");
        string resultPath = Path.Combine(outputDir, "ResultDocument.odt");

        // ---------- Create the main document with a placeholder ----------
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the main document.");
        mainBuilder.Writeln("PLACEHOLDER"); // The text to be replaced.
        mainBuilder.Writeln("End of the main document.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // ---------- Create the document that will replace the placeholder ----------
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is the inserted document.");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // ---------- Set up FindReplaceOptions with a custom callback ----------
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentHandler(insertDocPath)
        };

        // Perform the replace operation. The placeholder will be replaced by the whole document.
        mainDoc.Range.Replace(new Regex("PLACEHOLDER"), string.Empty, options);

        // ---------- Save the resulting document as ODT ----------
        mainDoc.Save(resultPath, SaveFormat.Odt);

        // ---------- Validate that the file was created ----------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The ODT file was not created.");

        // Clean up intermediate files (optional).
        // File.Delete(mainDocPath);
        // File.Delete(insertDocPath);
    }

    // Callback that inserts a document at the location of the match.
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

            // The match is inside a paragraph; insert the document after that paragraph.
            Paragraph para = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(para, subDoc);

            // Remove the placeholder paragraph.
            para.Remove();

            // Skip further processing for this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of docToInsert after the specified paragraph.
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
