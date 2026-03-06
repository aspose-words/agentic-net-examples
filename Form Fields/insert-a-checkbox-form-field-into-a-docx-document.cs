using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some explanatory text.
        builder.Write("Please tick the box: ");

        // Insert a checkbox form field.
        // Parameters: name, checked status (false = unchecked), size (0 = auto size).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // Example of setting an explicit size (optional).
        // checkBox.IsCheckBoxExactSize = true;
        // checkBox.CheckBoxSize = 12; // size in points

        // Save the document to a DOCX file.
        doc.Save("CheckboxFormField.docx");
    }
}
