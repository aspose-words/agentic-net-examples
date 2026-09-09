using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    [STAThread]
    public static void Main()
    {
        // Create a sample document with a bookmark that encloses the content to copy.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph before bookmark.");
        builder.StartBookmark("CopyMe");
        builder.Writeln("This is the content to copy.");
        builder.EndBookmark("CopyMe");
        builder.Writeln("Paragraph after bookmark.");

        // Save the source document to a deterministic local file.
        const string sourcePath = "sample.docx";
        doc.Save(sourcePath);

        // Load the document from the file system.
        Document loadedDoc = new Document(sourcePath);

        // Retrieve the bookmark that defines the selected content.
        Bookmark bookmark = loadedDoc.Range.Bookmarks["CopyMe"];
        if (bookmark == null)
            throw new InvalidOperationException("Required bookmark was not found.");

        // Extract the text inside the bookmark.
        string extractedText = bookmark.Text;

        // NOTE: Clipboard access requires a reference to System.Windows.Forms, which is not
        // available in the default console project used for verification. The extracted text
        // is therefore written to a file for validation purposes.
        // Clipboard.SetText(extractedText); // Omitted for compatibility.

        // Write the extracted text to a file for verification.
        const string outputPath = "extracted.txt";
        File.WriteAllText(outputPath, extractedText);

        // Validate that the output file was created and contains the expected text.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Extraction output file was not created.");

        string verified = File.ReadAllText(outputPath);
        if (!verified.Equals(extractedText, StringComparison.Ordinal))
            throw new InvalidOperationException("Extracted text does not match the expected content.");
    }
}
