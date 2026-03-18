using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup; // <-- added for StructuredDocumentTag and SdtType
using Aspose.Words.Saving;

namespace BatchHeaderContentControl
{
    class Program
    {
        static void Main()
        {
            // Folder containing the Word documents to process.
            string inputFolder = @"C:\Docs\Input";
            // Optional: folder to save processed files (can be the same as inputFolder to overwrite).
            string outputFolder = @"C:\Docs\Output";

            // Ensure the output folder exists.
            Directory.CreateDirectory(outputFolder);

            // Process each .docx file in the input folder.
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
                // The overload without the "isShowingPlaceholderText" parameter is available in
                // Aspose.Words 22.9 and later. If you target an older version, use the two‑parameter overload.
                StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);

                // Inside the content control, insert a DOCPROPERTY field that displays the document title.
                builder.InsertField(" DOCPROPERTY \"Title\" ");

                // Optionally, you can set a tag or title for the content control for identification.
                sdt.Title = "HeaderTitleControl";
                sdt.Tag = "HeaderTitle";

                // Update all fields in the document so the DOCPROPERTY reflects the metadata we set.
                doc.UpdateFields();

                // Save the modified document. Overwrite the original or write to the output folder.
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                doc.Save(outputPath);
            }
        }
    }
}
