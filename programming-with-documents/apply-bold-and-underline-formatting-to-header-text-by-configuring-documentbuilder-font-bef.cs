using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply bold and single underline formatting to the font.
        builder.Font.Bold = true;
        builder.Font.Underline = Underline.Single;

        // Insert the header text with the configured formatting.
        builder.Writeln("Header Text");

        // Save the document to a file.
        string outputPath = "FormattedHeader.docx";
        doc.Save(outputPath);

        // Indicate that the document has been saved.
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
