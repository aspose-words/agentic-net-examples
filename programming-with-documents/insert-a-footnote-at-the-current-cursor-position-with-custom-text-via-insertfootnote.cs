using System;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text to the document.
        builder.Writeln("This is a sample paragraph with a footnote reference.");

        // Insert a footnote at the current cursor position with custom text.
        builder.InsertFootnote(FootnoteType.Footnote, "This is the custom footnote text.");

        // Save the document.
        string outputPath = "FootnoteExample.docx";
        doc.Save(outputPath);
    }
}
