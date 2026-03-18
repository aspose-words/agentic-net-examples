using System;
using Aspose.Words;
using Aspose.Words.Fields;

class TextInputWithBookmark
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        // The first argument is the name of the form field.
        // Aspose.Words automatically creates a bookmark with the same name.
        FormField textInput = builder.InsertTextInput(
            name: "MyTextField",               // Form field (and bookmark) name
            type: TextFormFieldType.Regular,   // Regular text input
            format: "",                        // No specific format
            fieldValue: "Enter your name",     // Placeholder text
            maxLength: 0);                     // Unlimited length

        // Verify that the bookmark with the same name was created.
        Bookmark bookmark = doc.Range.Bookmarks["MyTextField"];
        if (bookmark != null)
        {
            Console.WriteLine($"Bookmark '{bookmark.Name}' was created automatically.");
        }

        // Save the document to disk.
        doc.Save("TextInputWithBookmark.docx");
    }
}
