using System;
using Aspose.Words;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optional: set the paragraph style to a heading style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Configure the font: make it bold and apply a single underline.
        builder.Font.Bold = true;
        builder.Font.Underline = Underline.Single;

        // Insert the header text. The formatting set above will be applied.
        builder.Writeln("Bold and Underlined Header");

        // Save the document to a file.
        doc.Save("HeaderBoldUnderline.docx");
    }
}
