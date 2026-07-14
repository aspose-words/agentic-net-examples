using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph alignment to justified.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;

        // Ensure word wrap is enabled (wrap by whole words).
        builder.ParagraphFormat.WordWrap = true;

        // Add a long paragraph to demonstrate justification and wrapping.
        string longText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                          "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                          "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi " +
                          "ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit " +
                          "in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur " +
                          "sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt " +
                          "mollit anim id est laborum.";
        builder.Writeln(longText);

        // Save the document to a file.
        doc.Save("JustifiedParagraph.docx");
    }
}
