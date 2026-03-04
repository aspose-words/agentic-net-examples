using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentConversion
{
    public class DocumentConversionExample
    {
        // Load a document from a file path and save it as PDF.
        public static void ConvertFileToPdf(string inputFilePath, string outputPdfPath)
        {
            // Load the document using the file‑based constructor.
            Document doc = new Document(inputFilePath);

            // Save the document; the .pdf extension automatically selects PDF format.
            doc.Save(outputPdfPath);
        }

        // Load a document from a stream and save it as PDF.
        public static void ConvertStreamToPdf(string inputFilePath, string outputPdfPath)
        {
            // Open a read‑only file stream.
            using (FileStream inputStream = File.OpenRead(inputFilePath))
            {
                // Load the document from the stream.
                Document doc = new Document(inputStream);

                // Save explicitly specifying the PDF format.
                doc.Save(outputPdfPath, SaveFormat.Pdf);
            }
        }

        // Load a document from a string (in‑memory) and save it as PDF.
        public static void ConvertStringToPdf(string documentContent, string outputPdfPath)
        {
            // Convert the string to a UTF‑8 byte array and wrap it in a MemoryStream.
            byte[] bytes = Encoding.UTF8.GetBytes(documentContent);
            using (MemoryStream memoryStream = new MemoryStream(bytes))
            {
                // Load the document from the memory stream.
                Document doc = new Document(memoryStream);

                // Create PDF‑specific save options via the provided factory method.
                SaveOptions pdfOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

                // Save the document using the created options.
                doc.Save(outputPdfPath, pdfOptions);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Adjust the paths below to point to real files on your machine.
            string inputPath = "input.docx";
            string outputPathFile = "output_file.pdf";
            string outputPathStream = "output_stream.pdf";
            string outputPathString = "output_string.pdf";

            // Example calls demonstrating each conversion method.
            DocumentConversionExample.ConvertFileToPdf(inputPath, outputPathFile);
            DocumentConversionExample.ConvertStreamToPdf(inputPath, outputPathStream);

            // For the string‑based conversion, read the source document into a string first.
            string content = File.ReadAllText(inputPath);
            DocumentConversionExample.ConvertStringToPdf(content, outputPathString);

            Console.WriteLine("All conversions completed successfully.");
        }
    }
}
