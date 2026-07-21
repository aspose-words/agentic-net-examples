using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create output folder and define file name
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);
        string outputFile = Path.Combine(outputFolder, "EndnoteExample.docx");

        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text to the document body
        builder.Write("This text will be referenced by an endnote.");

        // Insert an endnote with the specified reference text
        builder.InsertFootnote(FootnoteType.Endnote, "Endnote comment regarding referenced text.");

        // Save the document to the output path
        doc.Save(outputFile);
    }
}
