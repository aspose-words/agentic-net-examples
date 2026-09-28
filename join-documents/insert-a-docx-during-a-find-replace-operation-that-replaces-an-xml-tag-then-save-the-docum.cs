using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the sample files.
        string folderPath = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(folderPath);

        // Paths for the main document, the document to insert, and the final output.
        string mainDocPath = Path.Combine(folderPath, "MainDocument.docx");
        string insertDocPath = Path.Combine(folderPath, "InsertDocument.docx");
        string outputDocPath = Path.Combine(folderPath, "ResultDocument.docx");

        // -----------------------------------------------------------------
        // Create the main document containing a placeholder tag.
        // -----------------------------------------------------------------
        Document mainDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(mainDoc);
        builder.Writeln("This is the main document.");
        builder.Writeln("Here is the placeholder tag that will be replaced: <mytag>");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the document whose content will be inserted.
        // -----------------------------------------------------------------
        Document insertDoc = new Document();
        builder = new DocumentBuilder(insertDoc);
        builder.Writeln("=== Inserted Document Start ===");
        builder.Writeln("This content comes from the inserted DOCX file.");
        builder.Writeln("=== Inserted Document End ===");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents for processing.
        // -----------------------------------------------------------------
        Document srcMain = new Document(mainDocPath);
        Document srcInsert = new Document(insertDocPath);

        // -----------------------------------------------------------------
        // Perform find‑replace: replace the placeholder tag with the inserted document.
        // -----------------------------------------------------------------
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentHandler(srcInsert)
        };
        // Replace the placeholder text with an empty string; the callback will insert the document.
        srcMain.Range.Replace("<mytag>", string.Empty, options);

        // Save the resulting document as DOCX.
        srcMain.Save(outputDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Simple validation.
        // -----------------------------------------------------------------
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException("The result document was not saved correctly.");

        string resultText = new Document(outputDocPath).GetText();
        if (!resultText.Contains("Inserted Document Start"))
            throw new InvalidOperationException("The inserted document content was not found in the result.");
    }

    // Callback that inserts a document at the location of the found placeholder.
    private class InsertDocumentHandler : IReplacingCallback
    {
        private readonly Document _documentToInsert;

        public InsertDocumentHandler(Document documentToInsert)
        {
            _documentToInsert = documentToInsert ?? throw new ArgumentNullException(nameof(documentToInsert));
        }

        public ReplaceAction Replacing(ReplacingArgs e)
        {
            // Move a builder to the node that matched the placeholder.
            DocumentBuilder builder = new DocumentBuilder((Document)e.MatchNode.Document);
            builder.MoveTo(e.MatchNode);

            // Insert the whole document at this position.
            builder.InsertDocument(_documentToInsert, ImportFormatMode.KeepSourceFormatting);

            // Remove the placeholder node (the match node) so it does not appear in the result.
            e.MatchNode.Remove();

            // Skip the default replace action because we already handled insertion.
            return ReplaceAction.Skip;
        }
    }
}
