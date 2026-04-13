using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsFormProtectionDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph explaining the form.
            builder.Writeln("Please fill out the form fields below:");

            // Insert a text input form field.
            // Name: "TextField1", type: regular text, default value empty, placeholder text, max length 50.
            FormField textField = builder.InsertTextInput(
                "TextField1",
                TextFormFieldType.Regular,
                "",
                "Enter your name here",
                50);

            // Validate that the text field was created.
            if (textField == null || string.IsNullOrEmpty(textField.Name))
                throw new InvalidOperationException("Failed to create the text input form field.");

            // Insert a line break before the next field.
            builder.InsertBreak(BreakType.ParagraphBreak);

            // Insert a checkbox form field.
            // Name: "CheckBox1", unchecked by default, size 15 points.
            FormField checkBox = builder.InsertCheckBox("CheckBox1", false, 15);

            // Validate that the checkbox field was created.
            if (checkBox == null || string.IsNullOrEmpty(checkBox.Name))
                throw new InvalidOperationException("Failed to create the checkbox form field.");

            // Ensure at least one form field exists before proceeding.
            FormFieldCollection formFields = doc.Range.FormFields;
            if (formFields == null || formFields.Count == 0)
                throw new InvalidOperationException("The document does not contain any form fields.");

            // Protect the document so that only form fields can be edited.
            doc.Protect(ProtectionType.AllowOnlyFormFields);

            // Save the protected document.
            const string outputPath = "ProtectedFormFields.docx";
            doc.Save(outputPath);

            // Inform that the document was created (no interactive console required).
            Console.WriteLine($"Document saved to '{outputPath}' with form fields protection.");
        }
    }
}
