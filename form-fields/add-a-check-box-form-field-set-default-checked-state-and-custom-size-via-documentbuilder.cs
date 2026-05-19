using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some introductory text.
            builder.Write("Please check the box: ");

            // Insert a checkbox form field with a default checked state and a custom size.
            // Parameters: name, defaultValue, checkedValue, size (points).
            FormField checkBox = builder.InsertCheckBox("MyCheckBox", true, true, 30);

            // Enable exact size so that the custom size is applied.
            checkBox.IsCheckBoxExactSize = true;

            // Validate that the form field exists.
            FormField retrieved = doc.Range.FormFields["MyCheckBox"];
            if (retrieved == null)
                throw new InvalidOperationException("The checkbox form field was not found.");

            // Validate that the checkbox is checked as expected.
            if (!retrieved.Checked)
                throw new InvalidOperationException("The checkbox is not in the expected checked state.");

            // Save the document to disk.
            doc.Save("CheckBoxFormField.docx");
        }
    }
}
