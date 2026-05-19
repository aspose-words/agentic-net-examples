using System;
using Aspose.Words;
using Aspose.Words.Fields;   // Needed for FormField, TextFormFieldType, FieldType

namespace AsposeWordsRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert introductory text.
            builder.Writeln("Please fill in the form field below:");

            // Insert a text input form field.
            // Parameters: name, type, format, default text, max length.
            FormField textInput = builder.InsertTextInput(
                "MyTextInput",
                TextFormFieldType.Regular,
                "",
                "Initial value",
                50);

            // Save the original document (optional, just for reference).
            doc.Save("Original.docx");

            // ------------------------------------------------------------
            // Update the value of the text input form field located in a given range.
            // For this example, we use the range of the first section.
            // ------------------------------------------------------------
            Aspose.Words.Range targetRange = doc.FirstSection.Range;

            // The range's FormFields collection contains all form fields that intersect the range.
            if (targetRange.FormFields.Count > 0)
            {
                FormField fieldInRange = targetRange.FormFields[0];

                // Verify that we have the expected field.
                if (fieldInRange.Type == FieldType.FieldFormTextInput && fieldInRange.Name == "MyTextInput")
                {
                    // Update the field's result (the displayed value).
                    fieldInRange.Result = "Updated value";
                }
            }

            // Optionally update all fields in the document before saving.
            doc.UpdateFields();

            // Save the modified document.
            doc.Save("Updated.docx");
        }
    }
}
