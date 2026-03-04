using System;
using Aspose.Words;

namespace AsposeWordsCheckboxExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some explanatory text before the checkbox.
            builder.Writeln("Please check the box below:");

            // Insert a checkbox form field at the current cursor position.
            // Parameters: name, checkedValue (initial state), size (0 = auto size).
            builder.InsertCheckBox("MyCheckBox", false, 0);

            // Add a line break after the checkbox for readability.
            builder.Writeln();

            // Save the document in RTF format.
            doc.Save("CheckboxDocument.rtf");
        }
    }
}
