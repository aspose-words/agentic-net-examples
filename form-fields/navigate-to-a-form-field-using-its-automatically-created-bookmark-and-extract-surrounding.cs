using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to insert content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph before the form field.
        builder.Writeln("Paragraph before the form field.");

        // Insert a text input form field. A bookmark with the same name is created automatically.
        string fieldName = "MyTextField";
        builder.Write("Enter value: ");
        FormField textField = builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", "Default text", 0);
        builder.Writeln(" (end of line).");

        // Paragraph after the form field.
        builder.Writeln("Paragraph after the form field.");

        // Save the document.
        string filePath = "FormFieldBookmark.docx";
        doc.Save(filePath);

        // Load the document to simulate a separate read operation.
        Document loadedDoc = new Document(filePath);

        // Validate that the form field exists.
        FormField field = loadedDoc.Range.FormFields[fieldName];
        if (field == null)
            throw new InvalidOperationException($"Form field '{fieldName}' not found.");

        // Locate the automatically created bookmark.
        Bookmark bookmark = loadedDoc.Range.Bookmarks[fieldName];
        if (bookmark == null)
            throw new InvalidOperationException($"Bookmark for form field '{fieldName}' not found.");

        // The bookmark start node resides inside the paragraph that contains the form field.
        Paragraph containingParagraph = bookmark.BookmarkStart?.ParentNode as Paragraph;
        if (containingParagraph == null)
            throw new InvalidOperationException("Unable to locate the paragraph containing the form field.");

        // Extract the text of the containing paragraph.
        string paragraphText = containingParagraph.GetText();

        // Optionally extract surrounding paragraphs.
        string previousParagraphText = string.Empty;
        Paragraph previous = containingParagraph.PreviousSibling as Paragraph;
        if (previous != null)
            previousParagraphText = previous.GetText();

        string nextParagraphText = string.Empty;
        Paragraph next = containingParagraph.NextSibling as Paragraph;
        if (next != null)
            nextParagraphText = next.GetText();

        // Output the results.
        Console.WriteLine("Containing paragraph text:");
        Console.WriteLine(paragraphText.TrimEnd('\r', '\n'));

        Console.WriteLine("\nPrevious paragraph text:");
        Console.WriteLine(previousParagraphText.TrimEnd('\r', '\n'));

        Console.WriteLine("\nNext paragraph text:");
        Console.WriteLine(nextParagraphText.TrimEnd('\r', '\n'));
    }
}
