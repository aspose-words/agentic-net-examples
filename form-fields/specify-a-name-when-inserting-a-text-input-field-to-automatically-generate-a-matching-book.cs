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
        // A bookmark with the same name is automatically created.
        string fieldName = "CustomerName";
        builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", "Enter name", 0);

        // Verify that the bookmark was created.
        if (doc.Range.Bookmarks[fieldName] == null)
            throw new InvalidOperationException($"Bookmark '{fieldName}' was not created.");

        // Access the form field by its name and set a default value.
        FormField textField = doc.Range.FormFields[fieldName];
        if (textField == null)
            throw new InvalidOperationException($"Form field '{fieldName}' not found.");

        textField.Result = "John Doe";

        // Save the document to disk.
        doc.Save("FormFieldWithBookmark.docx");
    }
}
