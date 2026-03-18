using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Attempt to insert a text input form field with an empty name.
            InsertTextInputSafe(builder, string.Empty, TextFormFieldType.Regular, "", "Placeholder", 0);

            // Insert a valid form field for comparison.
            InsertTextInputSafe(builder, "ValidField", TextFormFieldType.Regular, "", "Enter text", 0);

            // Save the document.
            doc.Save("FormFields_Output.docx");
        }

        /// <summary>
        /// Inserts a text input form field, logging a warning if the name is empty.
        /// </summary>
        /// <param name="builder">The DocumentBuilder used for insertion.</param>
        /// <param name="name">The name of the form field.</param>
        /// <param name="type">The type of the text form field.</param>
        /// <param name="format">The format string for the field.</param>
        /// <param name="fieldValue">The default text shown in the field.</param>
        /// <param name="maxLength">Maximum length of the field (0 = unlimited).</param>
        private static void InsertTextInputSafe(DocumentBuilder builder, string name,
                                                TextFormFieldType type, string format,
                                                string fieldValue, int maxLength)
        {
            // Check for an empty name and log a warning if necessary.
            if (string.IsNullOrEmpty(name))
            {
                // Using Console.WriteLine for simplicity; replace with a proper logger if needed.
                Console.WriteLine("Warning: Attempted to insert a form field with an empty name. " +
                                  "A bookmark will not be created for this field.");
            }

            // Insert the form field using the Aspose.Words API.
            // The InsertTextInput method allows an empty string for the name, as documented.
            FormField field = builder.InsertTextInput(name, type, format, fieldValue, maxLength);

            // Optionally, you can set additional properties on the field here.
            // For example: field.CalculateOnExit = true;
        }
    }
}
