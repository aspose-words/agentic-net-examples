using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldSizeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a builder to add content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a checkbox form field with automatic size (size = 0).
            builder.Write("Sample checkbox: ");
            FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);
            // Enable explicit size handling.
            checkBox.IsCheckBoxExactSize = true;

            // Optional: save the initial document.
            doc.Save("InitialCheckBox.docx");

            // Retrieve the checkbox by its name from the form fields collection.
            FormField existingCheckBox = doc.Range.FormFields["MyCheckBox"];
            if (existingCheckBox == null)
                throw new InvalidOperationException("Checkbox form field not found.");

            // Change the size of the checkbox to 30 points.
            existingCheckBox.CheckBoxSize = 30.0;
            // Ensure the exact size flag remains true.
            existingCheckBox.IsCheckBoxExactSize = true;

            // Save the document with the updated checkbox size.
            doc.Save("ModifiedCheckBox.docx");
        }
    }
}
