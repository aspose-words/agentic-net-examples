using System;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set font to bold and underline before inserting the header text.
        builder.Font.Bold = true;
        builder.Font.Underline = Underline.Single;

        // Insert the header text.
        builder.Writeln("Bold and Underlined Header");

        // Optionally reset formatting for subsequent text.
        builder.Font.Bold = false;
        builder.Font.Underline = Underline.None;

        // Save the document.
        string outputPath = "HeaderFormatted.docx";
        doc.Save(outputPath);
    }
}
