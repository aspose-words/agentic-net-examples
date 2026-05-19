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

        // Insert a text input form field and give it a name.
        // A bookmark with the same name is created automatically.
        string fieldName = "UserName";
        FormField textField = builder.InsertTextInput(
            fieldName,                     // name of the form field (and bookmark)
            TextFormFieldType.Regular,    // type of the text field
            "",                           // format string (none)
            "Enter your name here",       // placeholder text
            0);                           // unlimited length

        // Validate that the form field exists in the collection.
        FormFieldCollection fields = doc.Range.FormFields;
        if (fields[fieldName] == null)
            throw new InvalidOperationException($"Form field '{fieldName}' was not found.");

        // Validate that the matching bookmark was created.
        if (doc.Range.Bookmarks[fieldName] == null)
            throw new InvalidOperationException($"Bookmark '{fieldName}' was not created.");

        // Optionally set an initial value for the field.
        textField.Result = "John Doe";

        // Save the document to disk.
        doc.Save("FormFieldWithBookmark.docx");
    }
}
