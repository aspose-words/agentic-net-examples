using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormProtectionExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to construct the document content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- Section 1 ----------
            // This section contains only regular text and will be read‑only.
            builder.Writeln("Section 1: This text is read‑only.");
            // Insert a section break to start a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // ---------- Section 2 ----------
            // This section will contain a form field that the user can edit.
            builder.Writeln("Section 2: Please fill in the form below.");
            builder.Write("Enter your name: ");

            // Insert a text input form field.
            FormField nameField = builder.InsertTextInput(
                "NameField",                     // Field name
                TextFormFieldType.Regular,       // Field type
                "",                              // Default text (empty)
                "Your name",                     // Placeholder text
                50);                             // Maximum length

            // Optionally set an initial value.
            nameField.Result = "John Doe";

            // Protect the entire document so that only form fields are editable.
            // All other content becomes read‑only.
            doc.Protect(ProtectionType.AllowOnlyFormFields);

            // Save the document to the output folder.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "FormProtected.docx");
            doc.Save(outputPath);

            // Inform the user where the file was saved (no interactive input required).
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
