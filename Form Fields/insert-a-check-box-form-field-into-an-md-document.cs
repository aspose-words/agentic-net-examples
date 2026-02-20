using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some explanatory text.
        builder.Writeln("Please tick the box below:");

        // Insert a check box form field.
        // Parameters: field name, initial checked state, size in points.
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);

        // Optional: configure additional properties.
        checkBox.IsCheckBoxExactSize = true;   // Use the exact size specified.
        checkBox.HelpText = "Click to toggle the checkbox";
        checkBox.OwnHelp = true;               // Use the custom help text.

        // Save the document.
        doc.Save("CheckBoxFormField.docx");
    }
}
