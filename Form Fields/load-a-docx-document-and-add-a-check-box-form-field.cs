using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the existing DOCX file.
            string inputPath = @"C:\Docs\InputDocument.docx";

            // Path where the modified document will be saved.
            string outputPath = @"C:\Docs\OutputDocument.docx";

            // Load the existing document.
            Document doc = new Document(inputPath);

            // Create a DocumentBuilder for the loaded document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert some explanatory text before the checkbox (optional).
            builder.Write("Please check the box: ");

            // Insert a checkbox form field.
            // Parameters: name, default checked value, size (0 = auto size).
            FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

            // Optionally set the checkbox to have an exact size.
            checkBox.IsCheckBoxExactSize = true;
            checkBox.CheckBoxSize = 12; // size in points

            // Save the modified document.
            doc.Save(outputPath);
        }
    }
}
