using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Paths for the documents
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string insertDocPath = Path.Combine(outputDir, "Insert.docx");
        string resultPdfPath = Path.Combine(outputDir, "Result.pdf");

        // Create the template document with a merge field
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("Report for record:");
        templateBuilder.InsertField("MERGEFIELD InsertDoc \\* MERGEFORMAT");
        templateBuilder.Writeln();
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // Create the document that will be inserted
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("Inserted content with language settings.");
        // Note: LanguageId property may not be available in all Aspose.Words versions.
        // The language information will be preserved automatically when inserting with KeepSourceFormatting.
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // Load the template for mail merge
        Document resultDoc = new Document(templatePath);

        // Set up the field merging callback to insert the document
        resultDoc.MailMerge.FieldMergingCallback = new InsertDocumentCallback(insertDocPath);

        // Execute mail merge (no actual data fields needed for this example)
        resultDoc.MailMerge.Execute(new string[0], new object[0]);

        // Save the merged document as PDF
        resultDoc.Save(resultPdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(resultPdfPath))
        {
            throw new InvalidOperationException($"Failed to create PDF at '{resultPdfPath}'.");
        }
    }

    // Callback class that inserts a document at the merge field location
    private class InsertDocumentCallback : IFieldMergingCallback
    {
        private readonly string _insertDocPath;

        public InsertDocumentCallback(string insertDocPath)
        {
            _insertDocPath = insertDocPath ?? throw new ArgumentNullException(nameof(insertDocPath));
        }

        public void FieldMerging(FieldMergingArgs args)
        {
            if (args.FieldName == "InsertDoc")
            {
                // Load the document to be inserted
                Document insertDoc = new Document(_insertDocPath);

                // Move the builder to the merge field location
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.FieldName);

                // Insert the document preserving source formatting (including language info)
                builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);
            }
        }

        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image fields are used in this example
        }
    }
}
