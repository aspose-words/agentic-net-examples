using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field at the current cursor position.
        // Parameters: name, type, format, placeholder text, maxLength (0 = unlimited).
        builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Enter your text here", 0);

        // Set up save options for MHTML output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Export the form field as an HTML <input> element (default behavior).
            ExportTextInputFormFieldAsText = false
        };

        // Save the document as an MHTML file.
        doc.Save("TextInputFormField.mhtml", saveOptions);
    }
}
