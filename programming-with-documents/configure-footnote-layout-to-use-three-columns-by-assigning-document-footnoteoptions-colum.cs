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

        // Use DocumentBuilder to add content and footnotes.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("This is a sample paragraph with a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "First footnote text.");

        builder.Writeln();
        builder.Write("Another paragraph with a second footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Second footnote text.");

        // Configure the footnote area to be displayed in three columns.
        doc.FootnoteOptions.Columns = 3;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "FootnotesThreeColumns.docx");
        doc.Save(outputPath);
    }
}
