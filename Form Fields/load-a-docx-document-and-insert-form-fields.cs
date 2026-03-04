using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsFormFieldsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the existing DOCX file that will be loaded.
            string inputPath = @"C:\Docs\InputDocument.docx";

            // Path where the modified document will be saved.
            string outputPath = @"C:\Docs\OutputDocument.docx";

            // Load the existing document.
            Document doc = new Document(inputPath);

            // Create a DocumentBuilder for the loaded document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the end of the document to start inserting form fields.
            builder.MoveToDocumentEnd();

            // Insert a paragraph break before adding the first form field.
            builder.Writeln();

            // Insert a checkbox form field.
            // Parameters: name, default checked state, size (in points).
            builder.InsertCheckBox("AcceptTerms", false, 50);

            // Insert a paragraph break between fields.
            builder.Writeln();

            // Insert a combo box (drop‑down) form field.
            // Parameters: name, list of items, selected index.
            string[] comboItems = { "Option A", "Option B", "Option C" };
            builder.InsertComboBox("SelectOption", comboItems, 0);

            // Insert a paragraph break between fields.
            builder.Writeln();

            // Insert a text input form field.
            // Parameters: name, type, default text, placeholder text, maximum length.
            builder.InsertTextInput("UserName", TextFormFieldType.Regular, "", "Enter your name", 30);

            // Save the modified document.
            doc.Save(outputPath);
        }
    }
}
