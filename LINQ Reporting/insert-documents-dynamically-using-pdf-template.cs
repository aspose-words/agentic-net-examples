using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDynamicInsert
{
    class Program
    {
        static void Main()
        {
            // Load the PDF template as a Word document.
            // The constructor loads the file and detects the format automatically.
            Document templateDoc = new Document("Template.pdf");

            // Create a DocumentBuilder to work with the loaded document.
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Move the cursor to a bookmark named "InsertHere" where the new content will be placed.
            // The bookmark must exist in the PDF template.
            builder.MoveToBookmark("InsertHere");

            // Load the document that should be inserted.
            Document docToInsert = new Document("ContentToInsert.docx");

            // Insert the whole source document at the current cursor position.
            // KeepSourceFormatting preserves the original formatting of the inserted document.
            builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

            // Optionally insert additional documents in the same way.
            // Document anotherDoc = new Document("Another.docx");
            // builder.InsertDocument(anotherDoc, ImportFormatMode.KeepSourceFormatting);

            // Prepare PDF save options (default options are sufficient for most cases).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the resulting document as a PDF file.
            // The overload with SaveOptions ensures the document is exported using the specified options.
            templateDoc.Save("Result.pdf", pdfOptions);
        }
    }
}
