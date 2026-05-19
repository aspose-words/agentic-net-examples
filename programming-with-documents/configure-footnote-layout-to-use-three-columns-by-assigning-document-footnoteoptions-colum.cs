using System;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some text and footnotes.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("This is a sample paragraph with a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "First footnote text.");

        builder.Writeln(); // Start a new paragraph.
        builder.Write("Another paragraph with a second footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Second footnote text.");

        // Configure the footnote area to be displayed in three columns.
        doc.FootnoteOptions.Columns = 3;

        // Save the document to the local file system.
        doc.Save("FootnoteColumns.docx");
    }
}
