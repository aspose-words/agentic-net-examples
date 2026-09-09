using System;
using System.Data;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;

namespace AsposeWordsMailMergeInsertDoc
{
    public class Program
    {
        public static void Main()
        {
            // Prepare a working folder.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "WorkFolder");
            Directory.CreateDirectory(workDir);

            // 1. Create a template document that contains a MERGEFIELD where the DOCX will be inserted.
            string templatePath = Path.Combine(workDir, "Template.docx");
            Document templateDoc = CreateTemplateDocument(templatePath);

            // 2. Create a sample document that will be inserted during mail merge.
            string insertDocPath = Path.Combine(workDir, "InsertDoc.docx");
            Document insertDoc = CreateInsertDocument(insertDocPath);

            // 3. Build a DataTable that supplies the path of the document to insert for each record.
            DataTable mailData = new DataTable("MailData");
            mailData.Columns.Add("DocPath", typeof(string));
            // Add three records – all pointing to the same sample document.
            mailData.Rows.Add(insertDocPath);
            mailData.Rows.Add(insertDocPath);
            mailData.Rows.Add(insertDocPath);

            // 4. Load the template and assign a custom callback that will replace the merge field
            //    with the content of the document referenced in the data row.
            Document mergedDoc = new Document(templatePath);
            mergedDoc.MailMerge.FieldMergingCallback = new InsertDocCallback();

            // 5. Execute the mail merge. The callback performs the insertion.
            mergedDoc.MailMerge.Execute(mailData);

            // 6. Save the result as PDF.
            string outputPdf = Path.Combine(workDir, "Result.pdf");
            mergedDoc.Save(outputPdf, SaveFormat.Pdf);

            // Simple validation – ensure the PDF was created.
            if (!File.Exists(outputPdf))
                throw new InvalidOperationException("The PDF output was not created.");
        }

        // Creates a minimal template with a single MERGEFIELD named "DocContent".
        private static Document CreateTemplateDocument(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Report Header");
            builder.InsertField(" MERGEFIELD DocContent ");
            builder.Writeln();
            builder.Writeln("Report Footer");

            doc.Save(filePath, SaveFormat.Docx);
            return doc;
        }

        // Creates a sample DOCX whose language information we want to keep.
        private static Document CreateInsertDocument(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Inserted Document Content");

            // Set the language/locale of the run (LCID for en‑US is 1033).
            builder.Font.LocaleId = new CultureInfo("en-US").LCID;

            doc.Save(filePath, SaveFormat.Docx);
            return doc;
        }

        // Callback that inserts a whole document at the position of the merge field.
        private class InsertDocCallback : IFieldMergingCallback
        {
            public void FieldMerging(FieldMergingArgs args)
            {
                if (args.DocumentFieldName == "DocContent" &&
                    args.FieldValue is string docPath &&
                    File.Exists(docPath))
                {
                    // Load the source document.
                    Document srcDoc = new Document(docPath);

                    // Move the builder to the merge field location.
                    DocumentBuilder builder = new DocumentBuilder(args.Document);
                    builder.MoveToMergeField(args.DocumentFieldName);

                    // Insert the source document while preserving its formatting (including language).
                    builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

                    // Suppress the default field text.
                    args.Text = string.Empty;
                }
            }

            public void ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // No image handling required for this example.
            }
        }
    }
}
