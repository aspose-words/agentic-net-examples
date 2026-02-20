using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Attach a DocumentBuilder to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optional: add some explanatory text.
        builder.Writeln("Please tick the box:");

        // Insert a check box form field.
        // Parameters: field name, default checked state, size in points.
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);
        // Ensure the size we set is used.
        checkBox.IsCheckBoxExactSize = true;

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
