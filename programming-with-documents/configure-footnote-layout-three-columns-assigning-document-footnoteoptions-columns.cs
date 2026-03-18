using Aspose.Words;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build some content with a footnote so the footnote area is visible.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph that will have a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote text that will be laid out in three columns.");

        // Set the footnote layout to use three columns.
        doc.FootnoteOptions.Columns = 3;

        // Save the resulting document.
        doc.Save("FootnoteThreeColumns.docx");
    }
}
