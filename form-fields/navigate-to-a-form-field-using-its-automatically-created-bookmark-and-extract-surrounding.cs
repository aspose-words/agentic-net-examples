using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Path for the temporary document.
        const string filePath = "FormFieldExample.docx";

        // -------------------------------------------------
        // 1. Create a new document and insert a text input form field.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some surrounding text.
        builder.Writeln("Paragraph before the form field.");

        // Insert a text input form field named "MyTextInput".
        // A bookmark with the same name is automatically created.
        builder.InsertTextInput(
            "MyTextInput",                     // field name (bookmark name)
            TextFormFieldType.Regular,         // field type
            "",                                // default text
            "Enter your name here",            // placeholder text
            50);                               // maximum length

        // Add more text after the field.
        builder.Writeln("Paragraph after the form field.");

        // Save the document to disk.
        doc.Save(filePath);

        // -------------------------------------------------
        // 2. Load the document and locate the form field via its bookmark.
        // -------------------------------------------------
        Document loadedDoc = new Document(filePath);

        // Validate that the expected form field exists.
        FormField formField = loadedDoc.Range.FormFields["MyTextInput"];
        if (formField == null)
            throw new InvalidOperationException("Form field 'MyTextInput' was not found.");

        // Retrieve the automatically created bookmark.
        Bookmark bookmark = loadedDoc.Range.Bookmarks["MyTextInput"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark for 'MyTextInput' was not found.");

        // Navigate to the paragraph that contains the bookmark.
        Paragraph paragraph = bookmark.BookmarkStart.GetAncestor(NodeType.Paragraph) as Paragraph;
        if (paragraph == null)
            throw new InvalidOperationException("Unable to locate the paragraph containing the bookmark.");

        // Extract the full text of the surrounding paragraph.
        string paragraphText = paragraph.GetText();

        // Output the extracted paragraph text.
        Console.WriteLine("Extracted paragraph text:");
        Console.WriteLine(paragraphText);
    }
}
