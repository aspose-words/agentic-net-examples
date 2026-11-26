using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph before the form field.
        builder.Writeln("Paragraph before the form field.");

        // Insert a text input form field. This automatically creates a bookmark with the same name.
        builder.InsertTextInput("MyField", TextFormFieldType.Regular, "", "Default value", 0);
        builder.Writeln(); // End the paragraph that contains the form field.

        // Insert a paragraph after the form field.
        builder.Writeln("Paragraph after the form field.");

        // Save the document.
        const string outputPath = "FormFieldExample.docx";
        doc.Save(outputPath);

        // Locate the bookmark that was created for the form field.
        Bookmark bookmark = doc.Range.Bookmarks["MyField"];
        if (bookmark == null)
        {
            throw new InvalidOperationException("Bookmark 'MyField' was not found.");
        }

        // The bookmark start node resides inside a paragraph.
        Node bookmarkStartNode = bookmark.BookmarkStart;
        Paragraph fieldParagraph = bookmarkStartNode.ParentNode as Paragraph;
        if (fieldParagraph == null)
        {
            throw new InvalidOperationException("Unable to locate the paragraph containing the form field.");
        }

        // Extract text from the paragraph that contains the form field.
        string fieldParagraphText = fieldParagraph.GetText();

        // Extract text from the preceding paragraph, if it exists.
        Paragraph previousParagraph = fieldParagraph.PreviousSibling as Paragraph;
        string previousParagraphText = previousParagraph != null ? previousParagraph.GetText() : "(none)";

        // Extract text from the following paragraph, if it exists.
        Paragraph nextParagraph = fieldParagraph.NextSibling as Paragraph;
        string nextParagraphText = nextParagraph != null ? nextParagraph.GetText() : "(none)";

        // Output the results.
        Console.WriteLine("Document saved to: " + outputPath);
        Console.WriteLine();
        Console.WriteLine("Previous paragraph text:");
        Console.WriteLine(previousParagraphText);
        Console.WriteLine("Paragraph containing the form field:");
        Console.WriteLine(fieldParagraphText);
        Console.WriteLine("Next paragraph text:");
        Console.WriteLine(nextParagraphText);
    }
}
