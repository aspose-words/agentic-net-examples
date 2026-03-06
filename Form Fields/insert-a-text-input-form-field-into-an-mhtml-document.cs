using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt before the form field.
        builder.Write("Please enter your name: ");

        // Insert a text input form field.
        // Parameters: name, type, format, placeholder text, maxLength (0 = unlimited).
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 0);

        // Configure save options to keep the form field as an HTML INPUT element.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportTextInputFormFieldAsText = false // default, but set explicitly for clarity.
        };

        // Save the document as an MHTML file.
        doc.Save("FormField.mhtml", saveOptions);
    }
}
