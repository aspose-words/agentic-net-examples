using System;
using Aspose.Words;
using Aspose.Words.Notes; // Required for FootnoteType enum

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph.
        builder.Writeln("This is a sample paragraph with an endnote reference.");

        // Insert an endnote with reference text.
        builder.InsertFootnote(FootnoteType.Endnote, "This is the endnote text.");

        // Save the document.
        string outputPath = "EndnoteExample.docx";
        doc.Save(outputPath);
    }
}
