using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content and a footnote.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("This is a sample paragraph with a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

        // Configure the footnote layout to use three columns.
        doc.FootnoteOptions.Columns = 3;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "FootnoteColumns.docx");
        doc.Save(outputPath);
    }
}
