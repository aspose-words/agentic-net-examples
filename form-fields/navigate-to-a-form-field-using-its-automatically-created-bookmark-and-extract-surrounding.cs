using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph before the form field.
        builder.Writeln("Paragraph before the form field.");

        // Insert a text input form field. A bookmark with the same name is created automatically.
        builder.InsertTextInput("MyTextField", TextFormFieldType.Regular, "", "Default value", 0);

        // Paragraph after the form field.
        builder.Writeln("Paragraph after the form field.");

        // Save the document (required by the rules).
        const string filePath = "FormFieldDoc.docx";
        doc.Save(filePath);

        // Load the document (simulating a separate operation).
        Document loadedDoc = new Document(filePath);

        // Retrieve the bookmark that was automatically created for the form field.
        Bookmark bookmark = loadedDoc.Range.Bookmarks["MyTextField"];
        if (bookmark == null)
        {
            throw new InvalidOperationException("Bookmark for the form field not found.");
        }

        // The bookmark start node resides inside the paragraph that contains the form field.
        Node bookmarkStart = bookmark.BookmarkStart;
        Paragraph containingParagraph = bookmarkStart.ParentNode as Paragraph;
        if (containingParagraph == null)
        {
            throw new InvalidOperationException("Unable to locate the containing paragraph.");
        }

        // Extract the text of the containing paragraph.
        string paragraphText = containingParagraph.GetText();

        // Output the extracted paragraph text.
        Console.WriteLine("Containing paragraph text:");
        Console.WriteLine(paragraphText.TrimEnd('\r', '\n'));

        // Optionally, also retrieve the previous and next paragraphs.
        Paragraph previousParagraph = containingParagraph.PreviousSibling as Paragraph;
        Paragraph nextParagraph = containingParagraph.NextSibling as Paragraph;

        if (previousParagraph != null)
        {
            Console.WriteLine("\nPrevious paragraph text:");
            Console.WriteLine(previousParagraph.GetText().TrimEnd('\r', '\n'));
        }

        if (nextParagraph != null)
        {
            Console.WriteLine("\nNext paragraph text:");
            Console.WriteLine(nextParagraph.GetText().TrimEnd('\r', '\n'));
        }
    }
}
