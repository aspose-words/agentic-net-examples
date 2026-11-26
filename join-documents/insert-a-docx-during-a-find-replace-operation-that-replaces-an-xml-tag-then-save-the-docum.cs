using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    // Callback that inserts the content of a document at each match location.
    private class InsertDocumentCallback : IReplacingCallback
    {
        private readonly Document _documentToInsert;

        public InsertDocumentCallback(Document documentToInsert)
        {
            _documentToInsert = documentToInsert ?? throw new ArgumentNullException(nameof(documentToInsert));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // Import nodes from the document to insert into the destination document.
            var importer = new NodeImporter(_documentToInsert, e.MatchNode.Document, ImportFormatMode.KeepSourceFormatting);

            // The match node is a Run inside a Paragraph that holds the placeholder tag.
            // We will insert the imported paragraphs after the whole placeholder paragraph.
            Paragraph placeholderParagraph = e.MatchNode.ParentNode as Paragraph;
            if (placeholderParagraph == null)
                return ReplaceAction.Skip; // safety check

            // Keep a reference to the original paragraph so we can remove it later.
            Paragraph originalParagraph = placeholderParagraph;

            // Insert each paragraph from the source document after the placeholder paragraph.
            foreach (Paragraph srcParagraph in _documentToInsert.FirstSection.Body.Paragraphs)
            {
                Node importedNode = importer.ImportNode(srcParagraph, true);
                // Insert after the current placeholderParagraph (which moves forward as we insert).
                placeholderParagraph.ParentNode.InsertAfter(importedNode, placeholderParagraph);
                // Update the reference so the next insertion goes after the newly inserted paragraph.
                placeholderParagraph = (Paragraph)importedNode;
            }

            // Remove the original placeholder paragraph.
            originalParagraph.Remove();

            // Skip the default replace action because we already handled insertion.
            return ReplaceAction.Skip;
        }
    }

    public static void Main()
    {
        // Paths for the sample files.
        string mainDocPath = "MainDocument.docx";
        string insertDocPath = "InsertDocument.docx";
        string outputPath = "ResultDocument.docx";

        // -----------------------------------------------------------------
        // Create the main document containing a placeholder XML tag.
        // -----------------------------------------------------------------
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the beginning of the main document.");
        // Placeholder XML tag to be replaced.
        mainBuilder.Writeln("<myTag>");
        mainBuilder.Writeln("This is the end of the main document.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the document that will be inserted.
        // -----------------------------------------------------------------
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Document Start ===");
        insertBuilder.Writeln("This content comes from the inserted document.");
        insertBuilder.Writeln("=== Inserted Document End ===");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents for processing.
        // -----------------------------------------------------------------
        Document srcMain = new Document(mainDocPath);
        Document srcInsert = new Document(insertDocPath);

        // -----------------------------------------------------------------
        // Perform find‑replace: replace the XML tag with the content of srcInsert.
        // -----------------------------------------------------------------
        var options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentCallback(srcInsert)
        };
        // The replacement string is empty because the callback handles insertion.
        srcMain.Range.Replace("<myTag>", string.Empty, options);

        // -----------------------------------------------------------------
        // Save the resulting document as DOCX.
        // -----------------------------------------------------------------
        srcMain.Save(outputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validate that the output file was created and contains the inserted text.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The output file '{outputPath}' was not created.");

        Document resultDoc = new Document(outputPath);
        bool containsInsertedText = resultDoc.GetText().Contains("Inserted Document Start");
        if (!containsInsertedText)
            throw new InvalidOperationException("The inserted document content was not found in the result.");
    }
}
