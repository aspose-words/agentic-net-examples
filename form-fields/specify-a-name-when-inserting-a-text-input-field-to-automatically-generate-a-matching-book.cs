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
        // Aspose.Words automatically creates a bookmark with the same name.
        string fieldName = "CustomerName";
        string defaultText = "Enter name";
        builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", defaultText, 0);

        // Verify that the bookmark with the same name exists.
        Bookmark bookmark = doc.Range.Bookmarks[fieldName];
        if (bookmark == null)
        {
            throw new InvalidOperationException($"Bookmark '{fieldName}' was not created.");
        }

        // Access the form field via the FormFields collection and update its value.
        FormField textField = doc.Range.FormFields[fieldName];
        if (textField == null)
        {
            throw new InvalidOperationException($"Form field '{fieldName}' was not found.");
        }
        textField.Result = "John Doe";

        // Save the document to disk.
        doc.Save("FormFieldWithBookmark.docx");
    }
}
