using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text before the checkbox.
        builder.Writeln("Please check the box below:");

        // Insert a checkbox form field.
        // Parameters: name, defaultValue, checkedValue, size (0 = automatic size).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, false, 0);

        // Set the checkbox to have an exact size (optional).
        checkBox.IsCheckBoxExactSize = true;
        checkBox.CheckBoxSize = 12; // size in points

        // Save the document to a DOCX file.
        doc.Save("CheckboxForm.docx");
    }
}
