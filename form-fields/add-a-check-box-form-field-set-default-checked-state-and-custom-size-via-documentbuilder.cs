using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some explanatory text.
        builder.Writeln("Please tick the box if you agree:");

        // Insert a checkbox form field.
        // Name: "AgreeCheckBox"
        // defaultValue: true (checked by default when the document is opened)
        // checkedValue: true (current state)
        // size: 30 points.
        FormField checkBox = builder.InsertCheckBox("AgreeCheckBox", true, true, 30);

        // Apply the exact size specified above.
        checkBox.IsCheckBoxExactSize = true;

        // Save the document to disk.
        doc.Save("CheckboxFormField.docx");
    }
}
