using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Replacing; // Needed for FindReplaceOptions

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths.
        string insertDocPath = Path.Combine(outputDir, "InsertDoc.docx");
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string resultPdfPath = Path.Combine(outputDir, "Combined.pdf");

        // -----------------------------------------------------------------
        // 1. Create a document that will be inserted for each data row.
        // -----------------------------------------------------------------
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is inserted content for row {0}.");
        insertDoc.Save(insertDocPath);

        // -----------------------------------------------------------------
        // 2. Create a mail‑merge template containing a MERGEFIELD that will
        //    be replaced by the inserted document.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(template);
        templateBuilder.Writeln("Start of merged document");
        templateBuilder.InsertField("MERGEFIELD Document");
        templateBuilder.Writeln("End of merged document");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Set up a FieldMergingCallback that inserts the document at the
        //    merge field location for each record.
        // -----------------------------------------------------------------
        template.MailMerge.FieldMergingCallback = new InsertDocumentCallback(insertDocPath);

        // -----------------------------------------------------------------
        // 4. Build a simple data source with three rows.
        // -----------------------------------------------------------------
        DataTable data = new DataTable("Data");
        data.Columns.Add("Dummy");
        data.Rows.Add("Row1");
        data.Rows.Add("Row2");
        data.Rows.Add("Row3");

        // -----------------------------------------------------------------
        // 5. Execute the mail merge – the callback will insert a document
        //    for each row, effectively joining them.
        // -----------------------------------------------------------------
        template.MailMerge.Execute(data);

        // -----------------------------------------------------------------
        // 6. Save the combined result as PDF.
        // -----------------------------------------------------------------
        template.Save(resultPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 7. Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPdfPath))
            throw new InvalidOperationException("Failed to create the combined PDF file.");
    }

    // -----------------------------------------------------------------
    // Callback implementation that inserts a document at the merge field.
    // -----------------------------------------------------------------
    private class InsertDocumentCallback : IFieldMergingCallback
    {
        private readonly string _insertDocPath;

        public InsertDocumentCallback(string insertDocPath)
        {
            _insertDocPath = insertDocPath;
        }

        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Only handle the specific merge field.
            if (!args.DocumentFieldName.Equals("Document", StringComparison.OrdinalIgnoreCase))
                return;

            // Load the document to be inserted.
            Document docToInsert = new Document(_insertDocPath);

            // Replace the placeholder with the current (1‑based) record number.
            int recordNumber = args.RecordIndex + 1;
            docToInsert.Range.Replace("{0}", recordNumber.ToString(), new FindReplaceOptions());

            // Move the cursor to the merge field location and remove the field.
            DocumentBuilder builder = new DocumentBuilder(args.Document);
            builder.MoveToMergeField(args.DocumentFieldName);

            // Insert the prepared document at the cursor position.
            builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

            // Prevent the default text insertion for this field.
            args.Text = string.Empty;
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
