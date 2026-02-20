using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // Retrieve the form field by its name (replace with your actual field name).
        FormField formField = doc.Range.FormFields["MyTextInput"];
        if (formField != null)
        {
            // Set the value (result) of the form field.
            formField.Result = "New value";
        }

        // Save the updated document.
        doc.Save("output.docx");
    }
}
