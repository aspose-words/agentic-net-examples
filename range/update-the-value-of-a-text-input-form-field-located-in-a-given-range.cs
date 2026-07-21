using System;
using Aspose.Words;
using Aspose.Words.Fields;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Words.Range;

namespace AsposeWordsRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph that will contain the form field.
            builder.Writeln("Please enter your name:");

            // Insert a text input form field (regular type) with a default value.
            // Parameters: name, type, format, default text, max length (0 = unlimited).
            builder.InsertTextInput("UserName", TextFormFieldType.Regular, "", "John Doe", 0);

            // Obtain the range of the first paragraph.
            Paragraph paragraph = doc.FirstSection.Body.Paragraphs[0];
            AsposeRange paragraphRange = paragraph.Range;

            // Locate the form field within the paragraph's range by name.
            FormField nameField = paragraphRange.FormFields["UserName"];
            if (nameField != null)
            {
                // Update the displayed text of the form field.
                nameField.Result = "Jane Smith";
            }

            // Save the modified document.
            doc.Save("UpdatedFormField.docx");
        }
    }
}
