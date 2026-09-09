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

        // Write some prompt text.
        builder.Write("Please enter your name: ");

        // Insert a text input form field with a specific name.
        // The name "UserName" will also create a bookmark with the same name.
        FormField textField = builder.InsertTextInput(
            "UserName",                     // name of the form field (and bookmark)
            TextFormFieldType.Regular,      // type of the text field
            "",                             // format string (none)
            "John Doe",                     // default placeholder text
            0);                             // max length (0 = unlimited)

        // Verify that the bookmark was automatically created.
        if (doc.Range.Bookmarks["UserName"] == null)
            throw new InvalidOperationException("Bookmark 'UserName' was not created.");

        // Optionally, set a value for the form field.
        textField.Result = "Alice";

        // Save the document to disk.
        doc.Save("FormFieldWithBookmark.docx");
    }
}
