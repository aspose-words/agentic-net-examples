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

        // Set the font to be bold and underlined.
        builder.Font.Bold = true;
        builder.Font.Underline = Underline.Single;

        // Insert the header text with the configured formatting.
        builder.Writeln("Bold and Underlined Header");

        // Save the document to a file.
        string outputFile = "HeaderBoldUnderline.docx";
        doc.Save(outputFile);
    }
}
