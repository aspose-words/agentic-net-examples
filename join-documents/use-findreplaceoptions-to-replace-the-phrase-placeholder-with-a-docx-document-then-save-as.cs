using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare temporary folder and file paths
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(tempFolder);

        string mainDocPath = Path.Combine(tempFolder, "main.docx");
        string insertDocPath = Path.Combine(tempFolder, "insert.docx");
        string outputPath = Path.Combine(tempFolder, "merged.odt");

        // Create the main document containing the placeholder
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the main document.");
        mainBuilder.Writeln("PLACEHOLDER");
        mainBuilder.Writeln("End of main document.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // Create the document that will replace the placeholder
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is the inserted document.");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // Load documents for processing
        Document main = new Document(mainDocPath);
        Document insert = new Document(insertDocPath);

        // Locate the paragraph that contains the placeholder text
        Paragraph? placeholderParagraph = null;
        NodeCollection paragraphs = main.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph paragraph in paragraphs)
        {
            // GetText() returns the paragraph text including a trailing paragraph break
            if (paragraph.GetText().Trim() == "PLACEHOLDER")
            {
                placeholderParagraph = paragraph;
                break;
            }
        }

        if (placeholderParagraph == null)
            throw new InvalidOperationException("Placeholder paragraph not found.");

        // Insert the whole insert document at the placeholder location
        DocumentBuilder builder = new DocumentBuilder(main);
        builder.MoveTo(placeholderParagraph);
        builder.InsertDocument(insert, ImportFormatMode.KeepSourceFormatting);

        // Remove the original placeholder paragraph
        placeholderParagraph.Remove();

        // Save the merged document as ODT
        main.Save(outputPath, SaveFormat.Odt);

        // Validate that the output file exists
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged ODT file was not created.");

        // Load the saved ODT to verify content
        Document result = new Document(outputPath);
        string resultText = result.GetText();

        if (!resultText.Contains("This is the inserted document.") ||
            !resultText.Contains("This is the main document.") ||
            !resultText.Contains("End of main document."))
        {
            throw new InvalidOperationException("The merged document does not contain expected content.");
        }

        // Clean up temporary files (optional)
        try
        {
            File.Delete(mainDocPath);
            File.Delete(insertDocPath);
            // Keep the output file for inspection; uncomment the line below to delete it as well.
            // File.Delete(outputPath);
            Directory.Delete(tempFolder, true);
        }
        catch
        {
            // Ignored – cleanup is best‑effort.
        }
    }
}
