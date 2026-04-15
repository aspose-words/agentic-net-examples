using System;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build the document content and add footnotes.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph with a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "First footnote.");

        builder.Writeln("Another paragraph with a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Second footnote.");

        // Configure the footnote area to be displayed in three columns.
        doc.FootnoteOptions.Columns = 3;

        // Save the document.
        string outputPath = "FootnoteColumns.docx";
        doc.Save(outputPath);
    }
}
