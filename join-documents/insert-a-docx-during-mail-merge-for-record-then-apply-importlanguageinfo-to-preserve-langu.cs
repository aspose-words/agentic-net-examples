using System;
using System.Data;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace JoinDocumentsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a temporary folder for sample files.
            string folder = Path.Combine(Directory.GetCurrentDirectory(), "TempDocs");
            Directory.CreateDirectory(folder);

            // Define file paths.
            string templatePath = Path.Combine(folder, "Template.docx");
            string insertPath = Path.Combine(folder, "Insert.docx");
            string outputPdfPath = Path.Combine(folder, "Result.pdf");

            // -----------------------------------------------------------------
            // Create the main template document containing a merge field.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Report Header");
            builder.InsertField("InsertDoc");
            builder.Writeln("Report Footer");
            templateDoc.Save(templatePath, SaveFormat.Docx);

            // -----------------------------------------------------------------
            // Create the document that will be inserted and set its language.
            // -----------------------------------------------------------------
            var insertDoc = new Document();
            var insertBuilder = new DocumentBuilder(insertDoc);
            // Set language to English (US) using LocaleId (LCID 1033).
            insertBuilder.Font.LocaleId = new CultureInfo("en-US").LCID;
            insertBuilder.Writeln("This is inserted content with English (US) language.");
            insertDoc.Save(insertPath, SaveFormat.Docx);

            // -----------------------------------------------------------------
            // Load the template for mail merge.
            // -----------------------------------------------------------------
            var mainDoc = new Document(templatePath);

            // Set up a custom callback to replace the merge field with the insert document.
            var handler = new InsertDocumentHandler(insertPath);
            mainDoc.MailMerge.FieldMergingCallback = handler;

            // Prepare a dummy data source (the field value itself is not used).
            var dataTable = new DataTable();
            dataTable.Columns.Add("InsertDoc");
            dataTable.Rows.Add(string.Empty);

            // Execute mail merge; the callback will insert the document.
            mainDoc.MailMerge.Execute(dataTable);

            // -----------------------------------------------------------------
            // Save the merged result as PDF.
            // -----------------------------------------------------------------
            mainDoc.Save(outputPdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(outputPdfPath))
                throw new Exception("Output PDF was not created.");

            // Simple validation that the inserted content exists in the final document.
            string mergedText = mainDoc.GetText();
            if (!mergedText.Contains("This is inserted content"))
                throw new Exception("Inserted document content not found in the merged result.");
        }

        // Custom field merging callback that inserts a document at the merge field location.
        private class InsertDocumentHandler : IFieldMergingCallback
        {
            private readonly string _docPath;

            public InsertDocumentHandler(string docPath)
            {
                _docPath = docPath;
            }

            public void FieldMerging(FieldMergingArgs args)
            {
                if (args.FieldName == "InsertDoc")
                {
                    // Load the document to be inserted.
                    var insertDoc = new Document(_docPath);

                    // Move the builder to the merge field position.
                    var builder = new DocumentBuilder(args.Document);
                    builder.MoveToMergeField(args.FieldName);

                    // Insert the document at the current position, keeping source formatting.
                    builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);
                }
                else
                {
                    // Default handling for other fields.
                    args.Text = args.FieldValue?.ToString() ?? string.Empty;
                }
            }

            public void ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // No image fields are used in this example.
            }
        }
    }
}
