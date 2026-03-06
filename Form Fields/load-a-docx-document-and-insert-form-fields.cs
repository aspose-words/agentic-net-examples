using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        // The Document constructor handles loading from a file path.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field.
        // Parameters: name, default checked state, size (points).
        builder.Writeln("Check the box:");
        builder.InsertCheckBox("MyCheckBox", false, 50);

        // Insert a combo box form field.
        // Parameters: name, list of items, selected index.
        builder.Writeln();
        builder.Writeln("Select an option:");
        string[] comboItems = { "Option A", "Option B", "Option C" };
        builder.InsertComboBox("MyComboBox", comboItems, 0);

        // Insert a text input form field.
        // Parameters: name, type, default text, placeholder, max length.
        builder.Writeln();
        builder.Writeln("Enter your name:");
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Enter name here", 30);

        // Save the modified document to a new file.
        doc.Save("Output.docx");
    }
}
