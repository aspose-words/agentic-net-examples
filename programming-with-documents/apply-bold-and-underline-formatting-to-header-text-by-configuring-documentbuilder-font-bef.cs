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

        // Set the paragraph style to Heading1 (optional, makes it a header).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Configure the font: make it bold and underlined.
        builder.Font.Bold = true;
        builder.Font.Underline = Underline.Single;

        // Insert the header text.
        builder.Writeln("Formatted Header Text");

        // Save the document to the local file system.
        string outputPath = "FormattedHeader.docx";
        doc.Save(outputPath);
    }
}
