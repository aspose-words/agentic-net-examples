using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertFormFieldsExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        // The Document constructor handles opening the file.
        Document doc = new Document("InputDocument.docx");

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document (optional, you can position it as needed).
        builder.MoveToDocumentEnd();

        // Insert a checkbox form field.
        // Parameters: name, default checked state, size (in points).
        builder.Writeln("Accept terms and conditions:");
        builder.InsertCheckBox("CheckBox_Terms", false, 50);
        builder.Writeln(); // Add a paragraph break.

        // Insert a combo box (drop‑down) form field.
        // Parameters: name, list of items, default selected index.
        string[] colors = { "Red", "Green", "Blue", "Yellow" };
        builder.Writeln("Select your favorite color:");
        builder.InsertComboBox("ComboBox_Color", colors, 0);
        builder.Writeln();

        // Insert a text input form field.
        // Parameters: name, type (regular text), default text, placeholder, maximum length.
        builder.Writeln("Enter your name:");
        builder.InsertTextInput("TextInput_Name", TextFormFieldType.Regular, "", "John Doe", 100);
        builder.Writeln();

        // Update all fields so that the form field results are reflected in the saved document.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save("OutputDocument.docx");
    }
}
