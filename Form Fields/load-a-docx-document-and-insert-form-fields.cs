using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field.
        // Parameters: name, default checked state, size (points).
        builder.InsertCheckBox("AcceptTerms", false, 50);
        builder.Writeln(); // Move to next line.

        // Insert a combo box form field.
        // Parameters: name, list of items, selected index.
        string[] colors = { "Red", "Green", "Blue", "Yellow" };
        builder.InsertComboBox("FavoriteColor", colors, 0);
        builder.Writeln();

        // Insert a text input form field.
        // Parameters: name, type, default text, placeholder, max length.
        builder.InsertTextInput("UserName", TextFormFieldType.Regular, "", "Enter your name", 30);
        builder.Writeln();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
