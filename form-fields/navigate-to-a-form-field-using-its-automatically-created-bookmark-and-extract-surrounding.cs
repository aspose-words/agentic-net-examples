using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        const string fileName = "FormFieldDemo.docx";

        // Create a new document and add a paragraph before the form field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the paragraph before the form field.");

        // Insert a text input form field. A bookmark with the same name ("MyTextField") is created automatically.
        builder.InsertTextInput("MyTextField", TextFormFieldType.Regular, "", "Enter value", 0);
        builder.Writeln("This is the paragraph after the form field.");

        // Save the document to disk.
        doc.Save(fileName);

        // Load the document back.
        Document loadedDoc = new Document(fileName);

        // Retrieve the bookmark that corresponds to the form field.
        Bookmark bookmark = loadedDoc.Range.Bookmarks["MyTextField"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'MyTextField' not found.");

        // The bookmark start node is inside the paragraph that contains the form field.
        // Its ParentNode is the Paragraph that we need.
        Paragraph paragraph = bookmark.BookmarkStart.ParentNode as Paragraph;
        if (paragraph == null)
            throw new InvalidOperationException("Parent paragraph of the bookmark not found.");

        // Extract the full text of the surrounding paragraph.
        string paragraphText = paragraph.GetText();

        // Output the extracted paragraph text.
        Console.WriteLine("Paragraph containing the form field:");
        Console.WriteLine(paragraphText);
    }
}
