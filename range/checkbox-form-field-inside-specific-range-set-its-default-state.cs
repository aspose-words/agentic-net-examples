using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text before the checkbox to define the insertion point.
        builder.Write("Please accept the terms: ");

        // Insert a checkbox form field at the current position.
        // Parameters: name, defaultValue (initial state), checkedValue (current state), size (0 = auto).
        FormField checkBox = builder.InsertCheckBox("AcceptTerms", true, true, 0);

        // The checkbox is now inserted with its default state set to checked.
        // Additional properties can be set if needed, e.g.:
        // checkBox.IsCheckBoxExactSize = true;
        // checkBox.CheckBoxSize = 12;

        // Save the document to a file.
        doc.Save("CheckboxFormField.docx");
    }
}
