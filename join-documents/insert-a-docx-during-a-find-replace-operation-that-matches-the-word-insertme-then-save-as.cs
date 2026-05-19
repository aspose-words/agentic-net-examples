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
        // Prepare a folder for temporary documents.
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsFolder);

        // Paths for the main document, the document to insert, and the final output.
        string mainDocPath = Path.Combine(docsFolder, "Main.docx");
        string insertDocPath = Path.Combine(docsFolder, "Insert.docx");
        string outputPath = Path.Combine(docsFolder, "Result.docx");

        // -----------------------------------------------------------------
        // Create the main document containing the placeholder "INSERTME".
        // -----------------------------------------------------------------
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the beginning of the main document.");
        mainBuilder.Writeln("INSERTME"); // Placeholder to be replaced.
        mainBuilder.Writeln("This is the end of the main document.");
        mainDoc.Save(mainDocPath);

        // ---------------------------------------------------------------
        // Create the document whose content will be inserted at the placeholder.
        // ---------------------------------------------------------------
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Document Start ===");
        insertBuilder.Writeln("Hello from the inserted document!");
        insertBuilder.Writeln("=== Inserted Document End ===");
        insertDoc.Save(insertDocPath);

        // ---------------------------------------------------------------
        // Load the main document and perform a find‑replace with a callback.
        // ---------------------------------------------------------------
        Document loadedMain = new Document(mainDocPath);
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentAtReplaceHandler(insertDocPath)
        };

        // Replace the placeholder text. The callback will insert the document.
        loadedMain.Range.Replace(new Regex("INSERTME"), string.Empty, options);

        // Save the resulting document.
        loadedMain.Save(outputPath, SaveFormat.Docx);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // (Optional) Clean up temporary files if desired.
        // File.Delete(mainDocPath);
        // File.Delete(insertDocPath);
    }

    // -----------------------------------------------------------------
    // Callback that inserts a document at the location of each match.
    // -----------------------------------------------------------------
    private class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        private readonly string _documentToInsertPath;

        public InsertDocumentAtReplaceHandler(string documentToInsertPath)
        {
            _documentToInsertPath = documentToInsertPath;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Load the document that will be inserted.
            Document subDoc = new Document(_documentToInsertPath);

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
                throw new ArgumentException("The destination node must be either a paragraph or a table.");

            CompositeNode dstStory = insertionDestination.ParentNode;

            NodeImporter importer = new NodeImporter(
                docToInsert,
                insertionDestination.Document,
                ImportFormatMode.KeepSourceFormatting);

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
