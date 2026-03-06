using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace InsertFormFieldsIntoDotx
{
    class Program
    {
        static void Main()
        {
            // Path to the folder where all files will be stored.
            string dataDir = Path.Combine(Environment.CurrentDirectory, "Data");
            Directory.CreateDirectory(dataDir);

            // -----------------------------------------------------------------
            // 1. Create a source document that contains various form fields.
            // -----------------------------------------------------------------
            Document sourceDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

            // Insert a check box form field.
            srcBuilder.Write("Accept terms: ");
            srcBuilder.InsertCheckBox("AcceptTerms", false, 50);
            srcBuilder.Writeln();

            // Insert a combo box (drop‑down) form field.
            srcBuilder.Write("Choose a color: ");
            srcBuilder.InsertComboBox("ColorChoice", new[] { "Red", "Green", "Blue" }, 0);
            srcBuilder.Writeln();

            // Insert a text input form field.
            srcBuilder.Write("Enter your name: ");
            srcBuilder.InsertTextInput("NameInput", TextFormFieldType.Regular, "", "John Doe", 100);
            srcBuilder.Writeln();

            // Save the source document to a temporary file (required for InsertDocument).
            string sourcePath = Path.Combine(dataDir, "SourceWithFormFields.docx");
            sourceDoc.Save(sourcePath, SaveFormat.Docx);

            // -----------------------------------------------------------------
            // 2. Load a DOTX template document.
            //    (Assume a template file named 'Template.dotx' exists in the same folder.)
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(dataDir, "Template.dotx");
            // For demonstration, create a minimal DOTX if it does not exist.
            if (!File.Exists(templatePath))
            {
                Document templateDoc = new Document();
                DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
                tmplBuilder.Writeln("=== Document Header ===");
                tmplBuilder.Writeln();
                tmplBuilder.Writeln("=== Content will be inserted below ===");
                tmplBuilder.Writeln();
                tmplBuilder.Writeln("=== Document Footer ===");
                templateDoc.Save(templatePath, SaveFormat.Dotx);
            }

            Document template = new Document(templatePath);
            DocumentBuilder builder = new DocumentBuilder(template);

            // Move the cursor to the position where the source content should be inserted.
            // Here we move to the end of the document, after the existing content.
            builder.MoveToDocumentEnd();

            // Insert a page break before the inserted content (optional).
            builder.InsertBreak(BreakType.PageBreak);

            // -----------------------------------------------------------------
            // 3. Insert the source document that contains form fields.
            // -----------------------------------------------------------------
            Document docToInsert = new Document(sourcePath);
            builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

            // -----------------------------------------------------------------
            // 4. Update fields so that form fields are correctly reflected.
            // -----------------------------------------------------------------
            template.UpdateFields();

            // -----------------------------------------------------------------
            // 5. Save the resulting document.
            // -----------------------------------------------------------------
            string resultPath = Path.Combine(dataDir, "ResultDocument.docx");
            template.Save(resultPath, SaveFormat.Docx);

            Console.WriteLine("Document created successfully at:");
            Console.WriteLine(resultPath);
        }
    }
}
