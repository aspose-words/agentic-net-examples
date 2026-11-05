using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

// Alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Words.Range;

namespace AsposeWordsRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to insert a paragraph with a text input form field.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Please fill the form below:");

            // Insert a text input form field named "MyTextInput" with default text "Old value".
            FormField textInput = builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Old value", 0);
            builder.Writeln(); // Add a new line after the form field.

            // The paragraph that currently contains the inserted form field.
            Paragraph paragraphWithField = builder.CurrentParagraph;

            // Obtain the range of that specific paragraph.
            AsposeRange paragraphRange = paragraphWithField.Range;

            // Retrieve the form field from the paragraph's range by its name.
            FormField fieldInRange = paragraphRange.FormFields["MyTextInput"];
            if (fieldInRange != null)
            {
                // Update the value of the text input form field.
                fieldInRange.Result = "New updated value";
            }

            // Save the modified document to a file.
            string outputPath = "UpdatedFormField.docx";
            doc.Save(outputPath);

            // Output the updated field value to the console for verification.
            Console.WriteLine($"Form field '{fieldInRange?.Name}' updated to: {fieldInRange?.Result}");
        }
    }
}
