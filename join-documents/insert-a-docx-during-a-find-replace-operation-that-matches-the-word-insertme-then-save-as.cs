using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for all sample files.
        string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "TempDocs");
        Directory.CreateDirectory(folderPath);

        // -------------------------------------------------
        // 1. Create the main document that contains the placeholder.
        // -------------------------------------------------
        string mainDocPath = Path.Combine(folderPath, "MainDocument.docx");
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the beginning of the main document.");
        mainBuilder.Writeln("Here is the placeholder: INSERTME");
        mainBuilder.Writeln("This is the end of the main document.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 2. Create the document whose content will be inserted.
        // -------------------------------------------------
        string insertDocPath = Path.Combine(folderPath, "InsertDocument.docx");
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Content Start ===");
        insertBuilder.Writeln("This paragraph comes from the inserted document.");
        insertBuilder.Writeln("=== Inserted Content End ===");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 3. Load both documents.
        // -------------------------------------------------
        Document destination = new Document(mainDocPath);
        Document sourceToInsert = new Document(insertDocPath);

        // -------------------------------------------------
        // 4. Replace the placeholder with the content of the source document.
        // -------------------------------------------------
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentCallback(sourceToInsert)
        };
        destination.Range.Replace("INSERTME", string.Empty, options);

        // -------------------------------------------------
        // 5. Save the resulting document.
        // -------------------------------------------------
        string resultPath = Path.Combine(folderPath, "ResultDocument.docx");
        destination.Save(resultPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 6. Validate that the result file exists and contains the inserted text.
        // -------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The result document was not created.");

        Document resultDoc = new Document(resultPath);
        if (!resultDoc.GetText().Contains("Inserted Content Start"))
            throw new InvalidOperationException("Inserted content was not found in the result document.");

        Console.WriteLine("Document merged successfully. Result saved at:");
        Console.WriteLine(resultPath);
    }

    // Callback that replaces the matched placeholder with the content of another document.
    private class InsertDocumentCallback : IReplacingCallback
    {
        private readonly Document _documentToInsert;

        public InsertDocumentCallback(Document documentToInsert)
        {
            _documentToInsert = documentToInsert ?? throw new ArgumentNullException(nameof(documentToInsert));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // The placeholder is inside a Run; get the containing Paragraph.
            Paragraph placeholderParagraph = e.MatchNode.ParentNode as Paragraph;
            if (placeholderParagraph == null)
                return ReplaceAction.Skip;

            // Prepare an importer to bring nodes from the source document into the destination.
            NodeImporter importer = new NodeImporter(_documentToInsert, placeholderParagraph.Document, ImportFormatMode.KeepSourceFormatting);

            // Insert each paragraph from the source document after the placeholder paragraph.
            Node referenceNode = placeholderParagraph;
            foreach (Paragraph srcParagraph in _documentToInsert.FirstSection.Body.Paragraphs)
            {
                Node importedParagraph = importer.ImportNode(srcParagraph, true);
                referenceNode.ParentNode.InsertAfter(importedParagraph, referenceNode);
                referenceNode = importedParagraph;
            }

            // Remove the original placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip the default replace action because we handled insertion manually.
            return ReplaceAction.Skip;
        }
    }
}
