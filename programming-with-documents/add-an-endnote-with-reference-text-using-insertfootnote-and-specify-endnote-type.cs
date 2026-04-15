using System;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some body text that will be referenced by the endnote.
        builder.Write("This is a paragraph with an endnote reference.");

        // Insert an endnote (FootnoteType.Endnote) with the desired reference text.
        builder.InsertFootnote(FootnoteType.Endnote, "This is the endnote text.");

        // Prepare an output folder and file path.
        string outputFolder = "Output";
        System.IO.Directory.CreateDirectory(outputFolder);
        string outputFile = System.IO.Path.Combine(outputFolder, "EndnoteExample.docx");

        // Save the document to disk.
        doc.Save(outputFile);
    }
}
