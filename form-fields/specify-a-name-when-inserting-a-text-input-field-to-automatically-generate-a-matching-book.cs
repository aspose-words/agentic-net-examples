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
        builder.Write("Enter your name: ");

        // Insert a text input form field with a name.
        // The name "UserName" will also create a bookmark with the same name.
        FormField textField = builder.InsertTextInput(
            "UserName",                     // name of the form field (and bookmark)
            TextFormFieldType.Regular,      // type of the text field
            "",                             // format string (none)
            "John Doe",                     // default placeholder text
            0);                             // unlimited length

        // Validate that the form field exists in the collection.
        FormFieldCollection formFields = doc.Range.FormFields;
        FormField retrievedField = formFields["UserName"];
        if (retrievedField == null)
            throw new InvalidOperationException("The form field 'UserName' was not found.");

        // Validate that the automatically created bookmark exists.
        Bookmark bookmark = doc.Range.Bookmarks["UserName"];
        if (bookmark == null)
            throw new InvalidOperationException("The bookmark 'UserName' was not created.");

        // Optionally, set a new value for the text field.
        retrievedField.Result = "Alice";

        // Save the document to the local file system.
        doc.Save("FormFieldBookmark.docx");
    }
}
