using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string sourceDocPath = @"C:\Docs\Template.docx";

            // Path where the resulting PDF will be saved.
            string outputPdfPath = @"C:\Docs\Result.pdf";

            // Load the existing DOCX document.
            Document doc = new Document(sourceDocPath);

            // Create a DocumentBuilder attached to the loaded document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Retrieve the image bytes from the database.
            // Replace this stub with actual database access logic.
            byte[] imageBytes = GetImageFromDatabase();

            // Insert the image at the current cursor position.
            // The overload that accepts a byte array inserts the image inline at 100% scale.
            builder.InsertImage(imageBytes);

            // Save the modified document as a PDF.
            doc.Save(outputPdfPath, SaveFormat.Pdf);
        }

        // Stub method representing image retrieval from a database.
        // In a real scenario, replace this with proper ADO.NET or ORM code.
        private static byte[] GetImageFromDatabase()
        {
            // Example: read an image file from disk to simulate database data.
            // In production, this would be a BLOB column read from the DB.
            string sampleImagePath = @"C:\Images\SampleImage.jpg";
            return File.ReadAllBytes(sampleImagePath);
        }
    }
}
