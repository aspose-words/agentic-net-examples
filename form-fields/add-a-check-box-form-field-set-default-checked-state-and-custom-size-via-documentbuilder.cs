using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Writeln("Document with a custom checkbox form field:");

        // Insert a checkbox form field.
        // Parameters: name, defaultValue (checked by default), checkedValue (current state), size in points.
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", true, true, 30);

        // Ensure the custom size is applied.
        checkBox.IsCheckBoxExactSize = true;
        checkBox.CheckBoxSize = 30; // size in points

        // Add a line break after the checkbox.
        builder.Writeln();

        // Save the document to the local file system.
        doc.Save("CheckBoxFormField.docx");
    }
}
