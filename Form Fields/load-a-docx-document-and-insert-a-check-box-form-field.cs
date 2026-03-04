using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertCheckBoxExample
{
    static void Main()
    {
        // Path to the folder that contains the input and output documents.
        string docsPath = @"C:\Docs\";

        // Load an existing DOCX file.
        Document doc = new Document(docsPath + "Input.docx");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document (or any desired position).
        builder.MoveToDocumentEnd();

        // Optional: add some explanatory text before the checkbox.
        builder.Write("Please check the box: ");

        // Insert a checkbox form field.
        // Parameters: name, default checked value, size (0 = auto size).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // If you want to specify an exact size, enable it and set the size in points.
        checkBox.IsCheckBoxExactSize = true;
        checkBox.CheckBoxSize = 12; // size in points

        // Save the modified document.
        doc.Save(docsPath + "Output.docx");
    }
}
