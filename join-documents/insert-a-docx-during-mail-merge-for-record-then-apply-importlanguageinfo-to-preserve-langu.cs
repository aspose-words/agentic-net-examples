using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a simple DOCX that will be inserted during mail merge.
        string insertDocPath = Path.Combine(outputDir, "InsertDoc.docx");
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is the inserted document.");
        // Save the source document that will be merged.
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // 2. Create a mail‑merge template containing a merge field where the DOCX will be placed.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document template = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(template);
        templateBuilder.Writeln("Report Header");
        templateBuilder.InsertField(" MERGEFIELD InsertDoc ");
        templateBuilder.Writeln("Report Footer");
        template.Save(templatePath, SaveFormat.Docx);

        // 3. Prepare a data source. The field value holds the path to the DOCX to insert.
        DataTable data = new DataTable("Data");
        data.Columns.Add("InsertDoc", typeof(string));
        data.Rows.Add(insertDocPath);

        // 4. Set up a field‑merging callback that loads the DOCX and inserts it.
        template.MailMerge.FieldMergingCallback = new InsertDocCallback();

        // 5. Execute mail merge.
        template.MailMerge.Execute(data);

        // 6. Save the merged document as PDF.
        string pdfPath = Path.Combine(outputDir, "MergedResult.pdf");
        template.Save(pdfPath, SaveFormat.Pdf);

        // 7. Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");
    }

    // Callback that replaces the merge field with the contents of a DOCX file.
    private class InsertDocCallback : IFieldMergingCallback
    {
        public void FieldMerging(FieldMergingArgs args)
        {
            // The field value is expected to be a file path to a DOCX document.
            if (args.FieldValue is string docPath && File.Exists(docPath))
            {
                // Load the document to be inserted.
                Document srcDoc = new Document(docPath);

                // Insert the document at the merge field location.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);
                builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

                // Prevent the default text insertion for this field.
                args.Text = string.Empty;
            }
        }

        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
