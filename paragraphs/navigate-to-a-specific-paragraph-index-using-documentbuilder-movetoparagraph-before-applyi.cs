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

        // Add several paragraphs to the document.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Move the builder's cursor to the second paragraph (index 1).
        // The second parameter (characterIndex) is set to 0 to position at the start of the paragraph.
        builder.MoveToParagraph(1, 0);

        // Apply formatting to the paragraph we have moved to.
        // Here we center-align the paragraph.
        builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Optionally, add additional text after formatting.
        builder.Writeln("This line is added after formatting the second paragraph.");

        // Save the document to the local file system.
        // The file will be created in the same directory as the executable.
        doc.Save("FormattedParagraphs.docx");
    }
}
