using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure the paragraph as a drop cap.
        // The drop cap will be the height of 3 lines of text.
        builder.ParagraphFormat.LinesToDrop = 3;
        // Position the drop cap inside the text margin.
        builder.ParagraphFormat.DropCapPosition = DropCapPosition.Normal;

        // Insert the drop cap character.
        builder.Writeln("D");

        // Return to normal paragraph formatting for the following text.
        builder.ParagraphFormat.LinesToDrop = 0;
        builder.ParagraphFormat.DropCapPosition = DropCapPosition.None;

        // Add regular text that will wrap around the drop cap.
        builder.Writeln("rop cap example text that wraps around the large letter.");

        // Save the document to the current directory.
        doc.Save("ParagraphDropCap.docx");
    }
}
