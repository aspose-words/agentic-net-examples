using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the main document that contains the placeholder text.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the main document.");
        mainBuilder.Writeln("PLACEHOLDER"); // This text will be replaced by another document.
        mainBuilder.Writeln("End of the main document.");

        // Create the document that will replace the placeholder.
        Document subDoc = new Document();
        DocumentBuilder subBuilder = new DocumentBuilder(subDoc);
        subBuilder.Writeln("This is the inserted document.");
        subBuilder.Writeln("It replaces the placeholder.");

        // Save the sub‑document to a temporary file so it can be loaded inside the callback.
        string subDocPath = Path.Combine(outputDir, "SubDocument.docx");
        subDoc.Save(subDocPath, SaveFormat.Docx);

        // Configure FindReplaceOptions with a custom callback that inserts the sub‑document.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentAtReplaceHandler(subDocPath)
        };

        // Perform the find‑and‑replace operation. The placeholder text is removed,
        // and the callback inserts the sub‑document at its location.
        mainDoc.Range.Replace(new Regex("\\bPLACEHOLDER\\b"), string.Empty, options);

        // Save the resulting document as ODT.
        string resultPath = Path.Combine(outputDir, "Result.odt");
        mainDoc.Save(resultPath, SaveFormat.Odt);

        // Simple validation to ensure the file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The ODT file was not created.");

        // The program finishes automatically; no user interaction is required.
    }

    // Callback class that inserts a document when a match is found.
    private class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        private readonly string _subDocumentPath;

        public InsertDocumentAtReplaceHandler(string subDocumentPath)
        {
            _subDocumentPath = subDocumentPath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document that will be inserted.
            Document subDoc = new Document(_subDocumentPath);

            // The match node is a Run inside a Paragraph; get the containing Paragraph.
            Paragraph para = (Paragraph)args.MatchNode.ParentNode;

            // Insert the sub‑document after the paragraph that held the placeholder.
            InsertDocument(para, subDoc);

            // Remove the original placeholder paragraph.
            para.Remove();

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }
    }

    // Inserts all nodes of a source document after a given paragraph or table.
    private static void InsertDocument(Node insertionDestination, Document docToInsert)
    {
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("The destination node must be a paragraph or a table.");

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
