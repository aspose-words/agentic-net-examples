using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

namespace BatchHeaderContentControl
{
    class Program
    {
        static void Main()
        {
            // Use folders relative to the executable so they always exist.
            string baseDir = AppContext.BaseDirectory;
            string inputFolder = Path.Combine(baseDir, "Input");
            string outputFolder = Path.Combine(baseDir, "Output");

            // Ensure the folders exist.
            Directory.CreateDirectory(inputFolder);
            Directory.CreateDirectory(outputFolder);

            // Process each .docx file in the input folder (if any).
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
            {
                // Load the document.
                Document doc = new Document(filePath);

                // Example: set some built‑in metadata based on the file name.
                string title = Path.GetFileNameWithoutExtension(filePath);
                doc.BuiltInDocumentProperties.Title = title;
                doc.BuiltInDocumentProperties.Author = Environment.UserName;
                doc.BuiltInDocumentProperties.Company = "My Company";

                // Create a DocumentBuilder for the loaded document.
                DocumentBuilder builder = new DocumentBuilder(doc);

                // Move the cursor to the primary header of the first section.
                builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

                // Insert a plain‑text content control (StructuredDocumentTag) into the header.
                StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);

                // Inside the content control, insert a DOCPROPERTY field that displays the document title.
                builder.InsertField(" DOCPROPERTY \"Title\" ");

                // Optionally, set a tag or title for the content control for identification.
                sdt.Title = "HeaderTitleControl";
                sdt.Tag = "HeaderTitle";

                // Update all fields in the document so the DOCPROPERTY reflects the metadata we set.
                doc.UpdateFields();

                // Save the modified document.
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                doc.Save(outputPath);
            }

            Console.WriteLine("Processing completed.");
        }
    }
}
