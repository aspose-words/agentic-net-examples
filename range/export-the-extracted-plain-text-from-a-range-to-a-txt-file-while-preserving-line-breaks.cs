using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample content with paragraph and page breaks.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third paragraph after a page break.");

        // Extract the plain text from the whole document range.
        // The returned string contains control characters such as '\r' for paragraph breaks
        // and '\f' for page breaks, which preserve the original layout.
        string extractedText = doc.Range.Text;

        // Define the output .txt file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ExtractedText.txt");

        // Write the extracted text to the file, preserving line breaks.
        // Use UTF8 encoding to support a wide range of characters.
        File.WriteAllText(outputPath, extractedText, Encoding.UTF8);

        // Optionally, demonstrate saving the document directly as plain text using TxtSaveOptions.
        // This also respects line breaks and can be used instead of manual file writing.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            // Ensure paragraph breaks are written as CRLF.
            ParagraphBreak = Environment.NewLine,
            // Preserve page breaks as form feed characters.
            ForcePageBreaks = true
        };
        string txtSavePath = Path.Combine(Environment.CurrentDirectory, "DocumentSavedAsTxt.txt");
        doc.Save(txtSavePath, txtOptions);
    }
}
