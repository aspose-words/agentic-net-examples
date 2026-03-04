using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text before the checkbox.
        builder.Write("Please tick the box: ");

        // Insert a checkbox form field at the current cursor position.
        // Parameters: name, defaultValue, checkedValue, size (0 = automatic size).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, false, 0);

        // Optional: set an explicit size for the checkbox.
        // checkBox.IsCheckBoxExactSize = true;
        // checkBox.CheckBoxSize = 12; // size in points

        // Add a paragraph break after the checkbox.
        builder.InsertParagraph();

        // Save the document to a DOCX file.
        doc.Save("CheckboxFormField.docx");
    }
}
