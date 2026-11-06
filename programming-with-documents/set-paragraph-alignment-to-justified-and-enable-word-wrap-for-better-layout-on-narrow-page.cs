using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph alignment to justified.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;

        // Enable word wrap for the paragraph (helps layout on narrow pages).
        builder.ParagraphFormat.WordWrap = true;

        // Add a long paragraph to demonstrate the formatting.
        builder.Writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
            "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. " +
            "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. " +
            "Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."
        );

        // Save the document to a file.
        const string outputPath = "JustifiedParagraph.docx";
        doc.Save(outputPath);
    }
}
