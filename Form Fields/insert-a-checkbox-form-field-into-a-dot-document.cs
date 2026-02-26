using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertCheckBoxIntoDot
{
    static void Main()
    {
        // Create a new blank document (will be saved as a DOT template).
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field at the current cursor position.
        // Parameters: name, checkedValue (initial state), size (0 = auto).
        builder.InsertCheckBox("MyCheckBox", false, 0);

        // Save the document as a Word template (.dot).
        doc.Save("Template.dot", SaveFormat.Dot);
    }
}
