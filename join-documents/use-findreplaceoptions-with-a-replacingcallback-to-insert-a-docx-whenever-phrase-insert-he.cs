using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    // Callback that inserts a document at the location of the found placeholder text.
    private class InsertDocumentHandler : IReplacingCallback
    {
        private readonly Document _documentToInsert;

        public InsertDocumentHandler(Document documentToInsert)
        {
            _documentToInsert = documentToInsert ?? throw new ArgumentNullException(nameof(documentToInsert));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // The placeholder is a Run node.
            if (e.MatchNode is not Run placeholderRun)
                return ReplaceAction.Skip;

            // Build on the destination document at the placeholder position.
            // Cast to Document because placeholderRun.Document returns DocumentBase.
            var builder = new DocumentBuilder((Document)placeholderRun.Document);
            builder.MoveTo(placeholderRun);
            // Insert the whole document.
            builder.InsertDocument(_documentToInsert, ImportFormatMode.KeepSourceFormatting);
            // Remove the placeholder run.
            placeholderRun.Remove();

            // Skip default replacement.
            return ReplaceAction.Skip;
        }
    }

    public static void Main()
    {
        // Paths for temporary files.
        string mainDocPath = "MainDocument.docx";
        string insertDocPath = "InsertDocument.docx";
        string resultDocPath = "Result.docx";

        // -----------------------------------------------------------------
        // Create the main document containing the placeholder.
        // -----------------------------------------------------------------
        var mainDoc = new Document();
        var mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the beginning of the document.");
        mainBuilder.Writeln("Please INSERT_HERE to add more content.");
        mainBuilder.Writeln("This is the end of the document.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the document that will be inserted.
        // -----------------------------------------------------------------
        var insertDoc = new Document();
        var insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Document Start ===");
        insertBuilder.Writeln("Here is the content that replaces the placeholder.");
        insertBuilder.Writeln("=== Inserted Document End ===");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load documents.
        // -----------------------------------------------------------------
        var destination = new Document(mainDocPath);
        var toInsert = new Document(insertDocPath);

        // -----------------------------------------------------------------
        // Set up FindReplaceOptions with the custom callback.
        // -----------------------------------------------------------------
        var options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentHandler(toInsert)
        };

        // Perform the find-and-replace.
        destination.Range.Replace("INSERT_HERE", string.Empty, options);

        // Save the resulting document.
        destination.Save(resultDocPath, SaveFormat.Docx);

        // Validate that the output file was created.
        if (!File.Exists(resultDocPath))
            throw new InvalidOperationException("Result document was not created.");

        // Clean up temporary files (keep the result for verification).
        File.Delete(mainDocPath);
        File.Delete(insertDocPath);
    }
}
