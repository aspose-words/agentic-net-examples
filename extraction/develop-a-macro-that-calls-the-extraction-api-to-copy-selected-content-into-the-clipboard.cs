using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Entry point of the console application.
    [STAThread]
    public static void Main()
    {
        // 1. Create a sample document with a bookmark that defines the content to copy.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Paragraph before the copy range.");
        builder.StartBookmark("CopyRange");
        builder.Writeln("This text will be copied to the clipboard.");
        builder.EndBookmark("CopyRange");
        builder.Writeln("Paragraph after the copy range.");

        const string samplePath = "sample.docx";
        sampleDoc.Save(samplePath);

        // 2. Load the document we just created.
        Document loadedDoc = new Document(samplePath);

        // 3. Locate the bookmark that bounds the content we want to extract.
        Bookmark copyBookmark = loadedDoc.Range.Bookmarks["CopyRange"];
        if (copyBookmark == null)
            throw new InvalidOperationException("Bookmark 'CopyRange' was not found in the document.");

        // 4. Extract the text inside the bookmark.
        string extractedText = copyBookmark.Text;
        if (string.IsNullOrEmpty(extractedText))
            throw new InvalidOperationException("No text was extracted from the bookmark.");

        // 5. Write the extracted text to a file (simulating the clipboard operation).
        const string outputPath = "clipboard.txt";
        File.WriteAllText(outputPath, extractedText);

        // 6. Verify that the output file was created and contains the expected text.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to write the extracted text to the output file.");

        string fileContent = File.ReadAllText(outputPath);
        if (fileContent != extractedText)
            throw new InvalidOperationException("Verification failed: file content does not match extracted text.");

        // Program completes without requiring user interaction.
    }
}
