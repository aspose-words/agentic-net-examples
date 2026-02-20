using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document from disk.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optional: add some explanatory text before the check box.
        builder.Writeln("Please check the box below:");

        // Insert a check box form field.
        // Parameters: name of the field, default checked state (false), size in points (50).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);

        // Configure additional properties of the check box.
        checkBox.IsCheckBoxExactSize = true;   // Use the exact size specified above.
        checkBox.HelpText = "Click to toggle the box";
        checkBox.OwnHelp = true;               // Use the custom help text.

        // Save the modified document to a new file.
        doc.Save("Output.docx");
    }
}
