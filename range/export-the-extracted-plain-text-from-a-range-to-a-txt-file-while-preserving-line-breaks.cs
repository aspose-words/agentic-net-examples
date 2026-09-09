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
        // The returned string contains control characters (e.g., \r for paragraph breaks,
        // \f for page breaks) which preserve the original line structure.
        string extractedText = doc.Range.Text;

        // Prepare an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Write the extracted text to a .txt file, preserving line breaks.
        string txtPath = Path.Combine(outputDir, "ExtractedText.txt");
        File.WriteAllText(txtPath, extractedText, Encoding.UTF8);

        // Optionally, demonstrate saving the document itself as plain text using Aspose.Words.
        // This is not required for the extraction task but shows the alternative approach.
        string txtSavePath = Path.Combine(outputDir, "DocumentSavedAsTxt.txt");
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            // Ensure paragraph breaks are written as CRLF.
            ParagraphBreak = Environment.NewLine,
            // Preserve page breaks as form feed characters.
            ForcePageBreaks = true
        };
        doc.Save(txtSavePath, saveOptions);
    }
}
