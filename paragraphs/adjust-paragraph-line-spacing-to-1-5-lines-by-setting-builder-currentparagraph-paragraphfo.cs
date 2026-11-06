using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with sample text.
        builder.Writeln("This is a sample paragraph. It will have line spacing set to 1.5 lines.");

        // Set the line spacing of the current paragraph to 1.5 lines.
        builder.CurrentParagraph.ParagraphFormat.LineSpacing = 1.5;

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
