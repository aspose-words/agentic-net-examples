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
        builder.Writeln("Please tick the box if you agree:");

        // Insert a checkbox form field at the current cursor position.
        // Parameters: name, defaultValue, checkedValue, size (0 = automatic size).
        FormField checkBox = builder.InsertCheckBox("AgreementCheckBox", false, false, 0);

        // Optionally, specify that the size should be calculated automatically.
        checkBox.IsCheckBoxExactSize = false;

        // Insert a new paragraph after the checkbox.
        builder.InsertParagraph();

        // Save the document to disk.
        doc.Save("CheckboxFormField.docx");
    }
}
