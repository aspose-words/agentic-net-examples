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

        // Insert a text input form field named "CustomerName".
        // This automatically creates a bookmark with the same name.
        builder.InsertTextInput("CustomerName", TextFormFieldType.Regular, "", "John Doe", 0);

        // Verify that the bookmark was created.
        Bookmark bookmark = doc.Range.Bookmarks["CustomerName"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'CustomerName' was not created.");

        // Access the form field by name and update its value.
        FormField formField = doc.Range.FormFields["CustomerName"];
        if (formField == null)
            throw new InvalidOperationException("Form field 'CustomerName' was not found.");

        formField.Result = "Jane Smith";

        // Save the document.
        doc.Save("FormFieldWithBookmark.docx");
    }
}
