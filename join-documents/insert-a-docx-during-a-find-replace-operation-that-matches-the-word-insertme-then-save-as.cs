using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(workFolder);

        // Paths for the documents.
        string mainDocPath = Path.Combine(workFolder, "Main.docx");
        string insertDocPath = Path.Combine(workFolder, "Insert.docx");
        string resultDocPath = Path.Combine(workFolder, "Result.docx");

        // -----------------------------------------------------------------
        // Create the main document containing the placeholder.
        // -----------------------------------------------------------------
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the main document.");
        mainBuilder.Writeln("INSERTME");
        mainBuilder.Writeln("End of the main document.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the document to be inserted.
        // -----------------------------------------------------------------
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Content Start ===");
        insertBuilder.Writeln("This content comes from the inserted document.");
        insertBuilder.Writeln("=== Inserted Content End ===");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the main document for processing and set up the replace.
        // -----------------------------------------------------------------
        Document srcDoc = new Document(mainDocPath);
        Document docToInsert = new Document(insertDocPath);

        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentCallback(docToInsert)
        };

        // Perform the replacement – the placeholder will be replaced by the whole document.
        srcDoc.Range.Replace("INSERTME", string.Empty, options);

        // Save the resulting document.
        srcDoc.Save(resultDocPath, SaveFormat.Docx);

        // Validate that the result file was created.
        if (!File.Exists(resultDocPath))
        {
            throw new InvalidOperationException("Result document was not created.");
        }
    }

    // -----------------------------------------------------------------
    // Callback that inserts a document at the location of the matched text.
    // -----------------------------------------------------------------
    private class InsertDocumentCallback : IReplacingCallback
    {
        private readonly Document _documentToInsert;

        public InsertDocumentCallback(Document documentToInsert)
        {
            _documentToInsert = documentToInsert ?? throw new ArgumentNullException(nameof(documentToInsert));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // The match node is a Run that contains the placeholder text.
            if (e.MatchNode is Run run && run.Document != null)
            {
                // Remove the placeholder text.
                run.Text = string.Empty;

                // Move a builder to the run's position and insert the document.
                DocumentBuilder builder = new DocumentBuilder(run.Document as Document);
                builder.MoveTo(run);
                builder.InsertDocument(_documentToInsert, ImportFormatMode.KeepSourceFormatting);
            }

            // Skip the default replace action because we have already handled insertion.
            return ReplaceAction.Skip;
        }
    }
}
