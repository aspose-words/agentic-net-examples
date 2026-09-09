using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some explanatory text.
            builder.Writeln("Please confirm the agreement:");

            // Insert a checkbox form field.
            // name: "AgreementCheckBox"
            // defaultValue: true (the default state when the document is opened)
            // checkedValue: true (the current state after insertion)
            // size: 30 points (custom size)
            FormField checkBox = builder.InsertCheckBox("AgreementCheckBox", true, true, 30);
            // Ensure the custom size is applied.
            checkBox.IsCheckBoxExactSize = true;

            // Add a line break after the checkbox.
            builder.Writeln();

            // Save the document.
            doc.Save("CheckboxFormField.docx");
        }
    }
}
