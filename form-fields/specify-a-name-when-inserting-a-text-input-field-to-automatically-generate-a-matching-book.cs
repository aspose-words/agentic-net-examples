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

        // Write a prompt before the form field.
        builder.Write("Please enter your name: ");

        // Insert a text input form field with a specific name.
        // The name "UserName" will also create a bookmark with the same name.
        FormField textField = builder.InsertTextInput(
            "UserName",                     // name of the form field (and bookmark)
            TextFormFieldType.Regular,      // type of the text field
            "",                             // format string (none)
            "John Doe",                     // default placeholder text
            0);                             // unlimited length

        // Verify that the form field exists in the collection.
        FormField? retrievedField = doc.Range.FormFields["UserName"];
        if (retrievedField == null)
            throw new InvalidOperationException("The form field 'UserName' was not found.");

        // Verify that the automatically created bookmark exists.
        Bookmark? bookmark = doc.Range.Bookmarks["UserName"];
        if (bookmark == null)
            throw new InvalidOperationException("The bookmark 'UserName' was not created.");

        // Optionally set a new value for the text field.
        retrievedField.Result = "Alice";

        // Save the document to disk.
        doc.Save("FormFieldWithBookmark.docx");
    }
}
