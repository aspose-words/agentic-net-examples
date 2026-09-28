using System;
using Aspose.Words;
using Aspose.Words.Notes; // Needed for the FootnoteType enum

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with a footnote reference.
        builder.Writeln("This paragraph contains a footnote reference.");

        // Insert a footnote using the correct enum reference.
        builder.InsertFootnote(FootnoteType.Footnote, "This is the footnote text.");

        // Configure footnote layout to use three columns.
        doc.FootnoteOptions.Columns = 3;

        // Save the document to disk.
        string outputPath = "FootnoteColumns.docx";
        doc.Save(outputPath);
    }
}
