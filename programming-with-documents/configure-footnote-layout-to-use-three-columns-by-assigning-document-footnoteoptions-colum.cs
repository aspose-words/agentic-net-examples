using System;
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
        builder.Writeln("This paragraph contains a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Sample footnote text.");

        // Configure the footnote area to be displayed in three columns.
        doc.FootnoteOptions.Columns = 3;

        // Save the document to the local file system.
        const string outputFile = "FootnoteColumns.docx";
        doc.Save(outputFile);
    }
}
