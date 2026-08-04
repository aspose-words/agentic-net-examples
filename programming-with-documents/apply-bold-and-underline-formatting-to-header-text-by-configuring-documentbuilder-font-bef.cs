using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure the font: make it bold and underline it.
        builder.Font.Bold = true;
        builder.Font.Underline = Underline.Single;

        // Insert the header text with the configured formatting.
        builder.Writeln("Header Text");

        // Save the document to a file in the current directory.
        doc.Save("HeaderFormatted.docx");
    }
}
