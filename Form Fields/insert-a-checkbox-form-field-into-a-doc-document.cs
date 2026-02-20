using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Writeln("Please tick the box below:");

        // Insert a check box form field.
        // Parameters: name, default checked state, size (points).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);

        // Optional: make the size exact and set help/status text.
        checkBox.IsCheckBoxExactSize = true;
        checkBox.HelpText = "Click to toggle the check box.";
        checkBox.OwnHelp = true;
        checkBox.StatusText = "Check box status.";
        checkBox.OwnStatus = true;

        // Save the document to disk.
        doc.Save("CheckboxFormField.docx");
    }
}
