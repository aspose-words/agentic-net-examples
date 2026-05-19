using System;
using System.IO;
using Aspose.Words;
using Newtonsoft.Json;

public class ExtractionMacro
{
    [STAThread]
    public static void Main()
    {
        // Create a sample document with a bookmark that encloses the content to copy.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Header text that will not be copied.");
        builder.StartBookmark("CopyMe");
        builder.Writeln("This paragraph will be copied to the clipboard.");
        builder.Writeln("Another line inside the bookmark.");
        builder.EndBookmark("CopyMe");
        builder.Writeln("Footer text that will not be copied.");

        doc.Save(docPath);

        // Load the document and locate the bookmark.
        Document loaded = new Document(docPath);
        Bookmark bookmark = loaded.Range.Bookmarks["CopyMe"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'CopyMe' was not found.");

        // Extract the text inside the bookmark.
        string extractedText = bookmark.Text;

        // Instead of using System.Windows.Forms.Clipboard (which may not be available),
        // write the extracted text to a temporary file that simulates the clipboard.
        const string simulatedClipboardPath = "clipboard.txt";
        File.WriteAllText(simulatedClipboardPath, extractedText);

        // Verify that the simulated clipboard now contains the expected text.
        string clipboardText = File.ReadAllText(simulatedClipboardPath);
        if (clipboardText != extractedText)
            throw new InvalidOperationException("Simulated clipboard content does not match the extracted text.");

        // Serialize the extracted text to a JSON file using Newtonsoft.Json.
        const string jsonPath = "extracted.json";
        string jsonPayload = JsonConvert.SerializeObject(new { Content = extractedText }, Formatting.Indented);
        File.WriteAllText(jsonPath, jsonPayload);

        // Validate that the JSON file was created.
        if (!File.Exists(jsonPath))
            throw new InvalidOperationException("Failed to create the JSON output file.");

        // Clean up temporary files (optional).
        // File.Delete(docPath);
        // File.Delete(jsonPath);
        // File.Delete(simulatedClipboardPath);
    }
}
