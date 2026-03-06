using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new DOCM document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a regular text input form field.
        // Parameters: name, type, format, placeholder text, maxLength (0 = unlimited).
        builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Enter your text here", 0);

        // Save the document. The .docm extension selects the macro‑enabled format.
        doc.Save("FormField.docm");
    }
}
