using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "JoinDocsExample");
        Directory.CreateDirectory(workFolder);

        // Paths for the sample documents.
        string mainDocPath = Path.Combine(workFolder, "main.docx");
        string insertDocPath = Path.Combine(workFolder, "insert.docx");
        string outputPath = Path.Combine(workFolder, "merged.odt");

        // Create the main document containing the placeholder.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the main document.");
        mainBuilder.Writeln("PLACEHOLDER");
        mainBuilder.Writeln("End of the main document.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // Create the document that will replace the placeholder.
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is the inserted document.");
        insertBuilder.Writeln("Additional inserted content.");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // Load the documents for processing.
        Document source = new Document(mainDocPath);
        Document toInsert = new Document(insertDocPath);

        // Replace the placeholder with the entire insert document using a custom callback.
        FindReplaceOptions replaceOptions = new FindReplaceOptions
        {
            ReplacingCallback = new ReplaceWithDocument(toInsert)
        };
        source.Range.Replace("PLACEHOLDER", string.Empty, replaceOptions);

        // Save the merged result as ODT.
        source.Save(outputPath, SaveFormat.Odt);

        // Validate that the output file exists and contains content from both documents.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged ODT file was not created.");

        Document result = new Document(outputPath);
        string resultText = result.GetText();

        if (!resultText.Contains("This is the main document.") ||
            !resultText.Contains("This is the inserted document.") ||
            !resultText.Contains("Additional inserted content.") ||
            !resultText.Contains("End of the main document."))
        {
            throw new InvalidOperationException("The merged document does not contain expected content.");
        }

        Console.WriteLine("Document merged and saved successfully to: " + outputPath);
    }

    // Custom callback that inserts a whole document at the match location.
    private class ReplaceWithDocument : IReplacingCallback
    {
        private readonly Document _documentToInsert;

        public ReplaceWithDocument(Document documentToInsert)
        {
            _documentToInsert = documentToInsert ?? throw new ArgumentNullException(nameof(documentToInsert));
        }

        public ReplaceAction Replacing(ReplacingArgs e)
        {
            // The node that contains the matched text (a Run node).
            Node matchNode = e.MatchNode;

            // Build a DocumentBuilder positioned at the match node.
            DocumentBuilder builder = new DocumentBuilder((Document)matchNode.Document);
            builder.MoveTo(matchNode);

            // Insert the whole document after the current position.
            builder.InsertDocument(_documentToInsert, ImportFormatMode.KeepSourceFormatting);

            // Remove the placeholder text node.
            matchNode.Remove();

            // Skip further processing for this match.
            return ReplaceAction.Skip;
        }
    }
}
