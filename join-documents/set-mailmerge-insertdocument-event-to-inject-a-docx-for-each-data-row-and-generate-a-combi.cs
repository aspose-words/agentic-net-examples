using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create sample documents that will be inserted during mail merge.
        string doc1Path = Path.Combine(outputDir, "InsertDoc1.docx");
        string doc2Path = Path.Combine(outputDir, "InsertDoc2.docx");
        CreateSampleDocument(doc1Path, "Content of the first inserted document.");
        CreateSampleDocument(doc2Path, "Content of the second inserted document.");

        // Create a mail‑merge template containing a merge field that will trigger document insertion.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("=== Begin of combined document ===");
        // The field name is "InsertDoc". The callback will replace this field with a whole document.
        builder.InsertField("MERGEFIELD InsertDoc");
        builder.Writeln("=== End of combined document ===");
        template.Save(templatePath);

        // Prepare data source – one row per document to be inserted.
        DataTable data = new DataTable("Data");
        data.Columns.Add("Dummy"); // Required column, not used.
        data.Rows.Add("Row1");
        data.Rows.Add("Row2");

        // List of documents to insert, aligned with the data rows.
        List<string> docsToInsert = new List<string> { doc1Path, doc2Path };

        // Set up a field‑merging callback that inserts the appropriate document.
        template.MailMerge.FieldMergingCallback = new InsertDocumentCallback(docsToInsert);

        // Execute mail merge – the callback will be invoked for each record.
        template.MailMerge.Execute(data);

        // Save the merged result as a single PDF file.
        string pdfPath = Path.Combine(outputDir, "Combined.pdf");
        template.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The combined PDF was not generated.");

        Console.WriteLine("Combined PDF created at: " + pdfPath);
    }

    // Helper method to create a simple one‑page DOCX with specified text.
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath, SaveFormat.Docx);
    }

    // Callback that replaces the merge field with a whole document.
    private class InsertDocumentCallback : IFieldMergingCallback
    {
        private readonly List<string> _documents;

        public InsertDocumentCallback(List<string> documents)
        {
            _documents = documents ?? throw new ArgumentNullException(nameof(documents));
        }

        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Determine which document to insert based on the current record index.
            int index = args.RecordIndex;
            if (index < 0 || index >= _documents.Count)
                throw new IndexOutOfRangeException("Record index out of range for document insertion.");

            string path = _documents[index];
            Document srcDoc = new Document(path);

            // Move the builder to the merge field and insert the whole document.
            DocumentBuilder builder = new DocumentBuilder(args.Document);
            builder.MoveToMergeField(args.FieldName);
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Suppress the original merge field text.
            args.Text = string.Empty;
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
