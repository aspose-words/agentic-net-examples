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

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text that will be referenced by the endnote.
        builder.Write("This sentence will have an endnote attached.");

        // Insert an endnote with the desired reference text.
        builder.InsertFootnote(FootnoteType.Endnote, "This is the endnote content.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EndnoteExample.docx");

        // Ensure the directory exists (in this case, the current directory always exists).
        doc.Save(outputPath);
    }
}
