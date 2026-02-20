using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some explanatory text before the check box.
        builder.Writeln("Please check the box:");

        // Insert a check box form field.
        // Parameters: field name, default checked state, size in points.
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);

        // Ensure the size is applied exactly as specified.
        checkBox.IsCheckBoxExactSize = true;

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
