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

        // Insert a text input form field with a specific name.
        // This automatically creates a bookmark with the same name.
        string fieldName = "CustomerName";
        string defaultText = "Enter name";
        builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", defaultText, 0);

        // Verify that the bookmark was created.
        if (doc.Range.Bookmarks[fieldName] == null)
        {
            throw new InvalidOperationException($"Bookmark '{fieldName}' was not created.");
        }

        // Access the form field by name and set a value.
        FormField? textField = doc.Range.FormFields[fieldName];
        if (textField != null)
        {
            textField.Result = "John Doe";
        }
        else
        {
            throw new InvalidOperationException($"Form field '{fieldName}' was not found.");
        }

        // Save the document to disk.
        doc.Save("FormFieldExample.docx");
    }
}
