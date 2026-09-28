using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some text.
        builder.Writeln("This is a sample paragraph whose line spacing will be set to 1.5 lines.");

        // Adjust the line spacing of the current paragraph to 1.5 lines.
        builder.CurrentParagraph.ParagraphFormat.LineSpacing = 1.5;

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
