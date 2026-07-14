using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field with a name.
        // A bookmark with the same name is created automatically.
        FormField textField = builder.InsertTextInput(
            "MyTextField",                     // name of the form field (and bookmark)
            TextFormFieldType.Regular,         // type of the text field
            "",                                // format string (none)
            "Enter your name here",            // placeholder text
            0);                                // unlimited length

        // Verify that the bookmark was created.
        Bookmark bookmark = doc.Range.Bookmarks["MyTextField"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark was not created.");

        // Optionally, set an initial value for the form field.
        textField.Result = "John Doe";

        // Save the document.
        doc.Save("FormFieldWithBookmark.docx");

        // Output confirmation.
        Console.WriteLine($"Form field '{textField.Name}' and bookmark '{bookmark.Name}' created successfully.");
    }
}
