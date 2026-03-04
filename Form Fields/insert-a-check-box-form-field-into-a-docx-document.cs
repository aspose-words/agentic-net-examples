using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertCheckBoxExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some explanatory text before the checkbox.
        builder.Write("Please tick the box if you agree: ");

        // Insert a checkbox form field.
        // Parameters: name, defaultValue, checkedValue, size (0 = auto size).
        FormField checkBox = builder.InsertCheckBox("AgreementCheckBox", false, false, 0);

        // Optionally set the exact size of the checkbox.
        checkBox.IsCheckBoxExactSize = true;
        checkBox.CheckBoxSize = 12; // size in points

        // Insert a paragraph break after the checkbox.
        builder.InsertParagraph();

        // Save the document to a DOCX file.
        doc.Save("CheckBoxFormField.docx");
    }
}
