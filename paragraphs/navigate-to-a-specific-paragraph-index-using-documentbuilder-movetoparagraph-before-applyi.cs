using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs to the document.
        builder.Writeln("Paragraph 0");
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");

        // Navigate to the third paragraph (index 2) and position the cursor at the start of the paragraph.
        // The second argument (nodeIndex) is set to 0 to place the cursor before the first node in the paragraph.
        builder.MoveToParagraph(2, 0);

        // Apply formatting to the paragraph: make the text bold and red.
        builder.Font.Bold = true;
        builder.Font.Color = Color.Red;

        // Insert a new line to demonstrate the applied formatting.
        builder.Writeln("This paragraph is now bold and red.");

        // Save the modified document to a file.
        doc.Save("Output.docx");
    }
}
