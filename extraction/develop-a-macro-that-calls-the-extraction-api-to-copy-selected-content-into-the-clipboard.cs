using System;
using System.IO;
using Aspose.Words;

public class ExtractionToClipboardExample
{
    // Entry point of the console application.
    [STAThread]
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document containing a bookmark that marks the
        //    range we want to extract.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Intro paragraph.");
        builder.StartBookmark("Selection");
        builder.Writeln("First line of selected content.");
        builder.Writeln("Second line of selected content.");
        builder.EndBookmark("Selection");
        builder.Writeln("Trailing paragraph.");

        // Save the sample document locally.
        const string sourcePath = "sample.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document from the file system.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Locate the bookmark that defines the selectable range.
        // -----------------------------------------------------------------
        Bookmark selectionBookmark = loadedDoc.Range.Bookmarks["Selection"];
        if (selectionBookmark == null)
            throw new InvalidOperationException("Bookmark 'Selection' was not found in the document.");

        // -----------------------------------------------------------------
        // 4. Extract the text inside the bookmark.
        // -----------------------------------------------------------------
        string extractedText = selectionBookmark.Text;
        if (string.IsNullOrEmpty(extractedText))
            throw new InvalidOperationException("No text was extracted from the bookmark.");

        // -----------------------------------------------------------------
        // 5. Copy the extracted text to the system clipboard.
        //    NOTE: System.Windows.Forms.Clipboard is not available in a
        //    plain console project without Windows Forms references.
        //    As an alternative, we write the text to a temporary file that
        //    can be opened manually, and we also output the text to the
        //    console for verification.
        // -----------------------------------------------------------------
        const string clipboardPlaceholderPath = "clipboard.txt";
        File.WriteAllText(clipboardPlaceholderPath, extractedText);
        Console.WriteLine("Extracted text (also saved to clipboard.txt):");
        Console.WriteLine(extractedText);

        // -----------------------------------------------------------------
        // 6. Write the extracted text to a verification file.
        // -----------------------------------------------------------------
        const string outputPath = "extracted.txt";
        File.WriteAllText(outputPath, extractedText);
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The extracted text file was not created.");

        // -----------------------------------------------------------------
        // 7. Indicate successful completion.
        // -----------------------------------------------------------------
        Console.WriteLine("Extraction completed successfully.");
    }
}
