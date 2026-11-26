using System;
using Aspose.Words;
using Aspose.Words.Notes; // Namespace containing the FootnoteType enum

public class EndnoteExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some sample text.
        builder.Writeln("This is a sample sentence with an endnote reference.");

        // Insert an endnote with the specified reference text.
        builder.InsertFootnote(FootnoteType.Endnote, "This is the endnote reference text.");

        // Define the output file path.
        string outputPath = "EndnoteExample.docx";

        // Save the document to the file system.
        doc.Save(outputPath);
    }
}
