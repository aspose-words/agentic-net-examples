using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

namespace AsposeWordsDemo
{
    class Program
    {
        static void Main()
        {
            // Define file paths.
            string docPath = Path.Combine(Environment.CurrentDirectory, "Sample.docx");
            string updatedDocPath = Path.Combine(Environment.CurrentDirectory, "Sample_Updated.docx");

            // -------------------------------------------------
            // 1. Create a new blank document (create rule).
            // -------------------------------------------------
            Document doc = new Document();

            // Use DocumentBuilder to insert some text.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, {Name}!");
            builder.Writeln("This document was generated on {Date}.");

            // Save the newly created document (save rule).
            doc.Save(docPath, SaveFormat.Docx);

            // -------------------------------------------------
            // 2. Load the existing document from file (load rule).
            // -------------------------------------------------
            Document loadedDoc = new Document(docPath);

            // Perform a find-and-replace operation on the document's range.
            // Replace placeholders with actual values.
            loadedDoc.Range.Replace("{Name}", "John Doe");
            loadedDoc.Range.Replace("{Date}", DateTime.Now.ToString("yyyy-MM-dd"));

            // Save the updated document.
            loadedDoc.Save(updatedDocPath, SaveFormat.Docx);

            // -------------------------------------------------
            // 3. Extract plain‑text representation of the updated document.
            // -------------------------------------------------
            PlainTextDocument plainText = new PlainTextDocument(updatedDocPath);
            string textContent = plainText.Text.Trim();

            // Output the extracted text to the console.
            Console.WriteLine("Extracted plain text:");
            Console.WriteLine(textContent);
        }
    }
}
