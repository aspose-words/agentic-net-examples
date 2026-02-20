using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text before the checkbox.
        builder.Writeln("Please tick the box if you agree:");

        // Insert a checkbox form field.
        // Parameters: field name, default checked state, size in points.
        FormField checkBox = builder.InsertCheckBox("AgreementCheckBox", false, 50);
        // Make the size exact as specified.
        checkBox.IsCheckBoxExactSize = true;

        // Optional: set help text displayed when the user presses F1.
        checkBox.HelpText = "Check to indicate agreement.";
        checkBox.OwnHelp = true;

        // Save the document in RTF format.
        doc.Save("CheckboxFormField.rtf");
    }
}
