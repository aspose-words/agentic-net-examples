using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Folder to store temporary documents.
        string folder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(folder);

        // Create two sample source DOCX files.
        string doc1Path = Path.Combine(folder, "Doc1.docx");
        string doc2Path = Path.Combine(folder, "Doc2.docx");
        CreateSampleDocument(doc1Path, "Content of Document 1.");
        CreateSampleDocument(doc2Path, "Content of Document 2.");

        // Create a template document that contains a MERGEFIELD which will be replaced by the inserted documents.
        string templatePath = Path.Combine(folder, "Template.docx");
        CreateTemplateDocument(templatePath);

        // Prepare a DataTable that holds the paths of the documents to be inserted.
        DataTable data = new DataTable("Docs");
        data.Columns.Add("DocPath", typeof(string));
        data.Rows.Add(doc1Path);
        data.Rows.Add(doc2Path);

        // Load the template.
        Document template = new Document(templatePath);

        // Register a field‑merging callback that will replace the MERGEFIELD with the corresponding document.
        template.MailMerge.FieldMergingCallback = new InsertDocumentCallback(data);

        // Execute the mail merge. The callback will insert the documents for each row.
        template.MailMerge.Execute(data);

        // Save the combined result as a PDF file.
        string outputPdf = Path.Combine(folder, "Combined.pdf");
        template.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException($"Failed to create PDF at '{outputPdf}'.");
    }

    // Creates a simple DOCX file with the specified text.
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath, SaveFormat.Docx);
    }

    // Creates a template DOCX that contains a MERGEFIELD named "Document".
    private static void CreateTemplateDocument(string filePath)
    {
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        // Insert a MERGEFIELD that will be replaced by the inserted document.
        builder.InsertField("MERGEFIELD Document");
        // Add a paragraph break after each inserted document for readability.
        builder.InsertParagraph();
        template.Save(filePath, SaveFormat.Docx);
    }

    // Callback that inserts a document at the position of the MERGEFIELD.
    private class InsertDocumentCallback : IFieldMergingCallback
    {
        private readonly DataTable _data;

        public InsertDocumentCallback(DataTable data)
        {
            _data = data ?? throw new ArgumentNullException(nameof(data));
        }

        public void FieldMerging(FieldMergingArgs args)
        {
            // Only handle the specific merge field.
            if (!string.Equals(args.DocumentFieldName, "Document", StringComparison.OrdinalIgnoreCase))
                return;

            // Get the path of the document for the current record.
            string path = _data.Rows[args.RecordIndex]["DocPath"].ToString();

            // Load the source document.
            Document srcDoc = new Document(path);

            // Move the cursor to the merge field and insert the document.
            DocumentBuilder builder = new DocumentBuilder(args.Document);
            builder.MoveToMergeField(args.DocumentFieldName);
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Suppress the default text insertion for this field.
            args.Text = string.Empty;
        }

        // Image fields are not used in this example.
        public void ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
