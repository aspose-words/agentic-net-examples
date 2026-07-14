using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add initial content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Paragraph before the checkbox.");

            // Insert a new paragraph that will contain the checkbox.
            builder.Writeln("Paragraph with checkbox:");

            // Move the builder cursor to the newly created paragraph.
            // The paragraph we just added is the last paragraph in the document body.
            Paragraph targetParagraph = doc.FirstSection.Body.Paragraphs[doc.FirstSection.Body.Paragraphs.Count - 1];
            builder.MoveTo(targetParagraph);

            // Insert a checkbox form field at the current position.
            // Name: "MyCheckBox", default checked state: true, size: 0 (auto).
            FormField checkBox = builder.InsertCheckBox("MyCheckBox", true, 0);

            // Set the default value of the checkbox (the state it will have when the document is opened).
            checkBox.Default = true;

            // Optionally set the current checked state (can be different from the default).
            checkBox.Checked = true;

            // Add another paragraph after the checkbox for clarity.
            builder.Writeln();
            builder.Writeln("Paragraph after the checkbox.");

            // Save the document to a file in the current directory.
            string outputPath = "CheckboxInRange.docx";
            doc.Save(outputPath);

            // Indicate completion (no interactive input required).
            Console.WriteLine($"Document saved to '{outputPath}'.");
        }
    }
}
