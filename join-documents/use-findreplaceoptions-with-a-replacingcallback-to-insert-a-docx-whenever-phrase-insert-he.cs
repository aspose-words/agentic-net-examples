using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create the main document with a placeholder.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the main document.");
        mainBuilder.Writeln("INSERT_HERE");
        mainBuilder.Writeln("End of the main document.");
        mainDoc.Save("main.docx", SaveFormat.Docx);

        // Create the document that will be inserted.
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is the inserted document.");
        insertBuilder.Writeln("It has multiple paragraphs.");
        insertDoc.Save("insert.docx", SaveFormat.Docx);

        // Load the main document for processing.
        Document doc = new Document("main.docx");

        // Set up FindReplaceOptions with a custom callback that inserts the document.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentCallback(insertDoc)
        };

        // Perform the find-and-replace; the replacement string is empty because the callback handles insertion.
        doc.Range.Replace("INSERT_HERE", string.Empty, options);

        // Save the resulting document.
        const string outputPath = "output.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the output file was created and contains inserted content.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        Document result = new Document(outputPath);
        string resultText = result.GetText();
        if (!resultText.Contains("This is the inserted document.") ||
            !resultText.Contains("It has multiple paragraphs."))
            throw new Exception("The inserted content was not found in the output document.");
    }
}

// Callback that replaces the placeholder with the contents of another document.
public class InsertDocumentCallback : IReplacingCallback
{
    private readonly Document _insertDoc;

    public InsertDocumentCallback(Document insertDoc)
    {
        _insertDoc = insertDoc ?? throw new ArgumentNullException(nameof(insertDoc));
    }

    public ReplaceAction Replacing(ReplacingArgs e)
    {
        // The match node for a text replace is a Run, not a Paragraph.
        Run? run = e.MatchNode as Run;
        if (run == null)
            return ReplaceAction.Skip;

        // Get the paragraph that contains the placeholder run.
        Paragraph placeholderParagraph = run.ParentParagraph;

        // Import nodes from the insert document into the main document.
        NodeImporter importer = new NodeImporter(_insertDoc, placeholderParagraph.Document, ImportFormatMode.KeepSourceFormatting);

        // Insert each node from the first section's body after the placeholder paragraph.
        foreach (Node srcNode in _insertDoc.FirstSection.Body)
        {
            Node importedNode = importer.ImportNode(srcNode, true);
            placeholderParagraph.ParentNode?.InsertAfter(importedNode, placeholderParagraph);
        }

        // Remove the placeholder paragraph.
        placeholderParagraph.Remove();

        // Skip the default replacement.
        return ReplaceAction.Skip;
    }
}
