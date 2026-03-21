using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace OleProgIdValidationExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The ProgID we intend to use for the OLE object.
            string progId = "Excel.Sheet.12";

            // Validate the ProgID before insertion.
            ValidateProgId(progId);

            // Create a dummy OLE data stream (e.g., a minimal ZIP header).
            byte[] dummyData = new byte[] { 0x50, 0x4B, 0x03, 0x04 };
            using (MemoryStream oleStream = new MemoryStream(dummyData))
            {
                // Insert as a normal (non‑icon) object; presentation stream is null.
                builder.InsertOleObject(oleStream, progId, false, null);
            }

            // Save the resulting document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleWithValidatedProgId.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }

        /// <summary>
        /// Ensures that the provided ProgID is not null, empty, or whitespace.
        /// Throws an ArgumentException if the validation fails.
        /// </summary>
        /// <param name="progId">The ProgID to validate.</param>
        static void ValidateProgId(string progId)
        {
            if (string.IsNullOrWhiteSpace(progId))
                throw new ArgumentException("ProgID cannot be null, empty, or whitespace.", nameof(progId));

            // Additional validation logic can be added here, e.g., checking against a whitelist.
        }
    }
}
