using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document that will serve as a DOT template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some explanatory text.
        builder.Writeln("Please check the box below:");

        // Insert a checkbox form field.
        // Parameters: field name, default checked state, size in points.
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);
        // Make the size exact rather than automatic.
        checkBox.IsCheckBoxExactSize = true;
        // Optional: provide help text that appears when the user presses F1.
        checkBox.HelpText = "Click to toggle the checkbox";
        checkBox.OwnHelp = true;

        // Save the document as a Word template (DOT format).
        doc.Save("CheckboxTemplate.dot", SaveFormat.Dot);
    }
}
