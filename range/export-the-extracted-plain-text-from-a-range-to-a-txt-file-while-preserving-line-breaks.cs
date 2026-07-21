using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportRangeToText
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample content with line breaks.
        builder.Writeln("First line.");
        builder.Writeln("Second line.");
        builder.Writeln("Third line.");

        // Extract the plain text of the whole document range.
        // The returned string contains control characters (CR) that represent line breaks.
        string extractedText = doc.Range.Text;

        // Define the output .txt file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedText.txt");

        // Write the extracted text to the file, preserving the original line breaks.
        File.WriteAllText(outputPath, extractedText);

        // Optionally, also demonstrate saving the document directly as plain text using TxtSaveOptions.
        // This shows an alternative way to obtain the same result.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            // Ensure paragraph breaks are written as CRLF (default) to keep line breaks.
            ParagraphBreak = Environment.NewLine
        };
        string directSavePath = Path.Combine(Directory.GetCurrentDirectory(), "DirectSave.txt");
        doc.Save(directSavePath, txtOptions);
    }
}
