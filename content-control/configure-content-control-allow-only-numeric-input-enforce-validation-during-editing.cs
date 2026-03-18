using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text form field that accepts only numbers.
        // TextFormFieldType.Number enforces numeric input.
        // The placeholder text "0" will be shown when the field is empty.
        FormField numericField = builder.InsertTextInput(
            name: "NumberField",
            type: TextFormFieldType.Number,
            format: "",               // No specific display format.
            fieldValue: "0",          // Placeholder/default value.
            maxLength: 0);            // Zero = no length limit.

        // Optional: set a simple numeric format (e.g., no decimals).
        numericField.TextInputFormat = "0";

        // Protect the document so that only form fields can be edited.
        // This ensures the numeric validation is enforced during editing.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the resulting document.
        doc.Save("NumericContentControl.docx");
    }
}
