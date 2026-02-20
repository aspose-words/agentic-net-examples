using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a line of text before the checkbox.
        builder.Writeln("Please tick the box below:");

        // Insert a check box form field.
        // Parameters: name, default checked state (false), size in points (50).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);
        // Ensure the size is applied exactly.
        checkBox.IsCheckBoxExactSize = true;
        // Optional: provide help text that appears on F1.
        checkBox.HelpText = "Click to toggle the checkbox";
        checkBox.OwnHelp = true;

        // Save the document to a DOCX file.
        doc.Save("CheckBoxFormField.docx");
    }
}
