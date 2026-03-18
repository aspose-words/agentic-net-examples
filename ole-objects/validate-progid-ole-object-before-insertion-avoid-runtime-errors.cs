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

            // Path to the file that will be embedded as an OLE object.
            string oleFilePath = @"C:\Data\Sample.xlsx";

            // The ProgID we intend to use for the OLE object.
            string progId = "Excel.Sheet.12";

            // Validate the ProgID before insertion.
            ValidateProgId(progId);

            // Insert the OLE object using the validated ProgID.
            using (FileStream oleStream = File.Open(oleFilePath, FileMode.Open, FileAccess.Read))
            {
                // Insert as a normal (non‑icon) object; presentation stream is null.
                builder.InsertOleObject(oleStream, progId, false, null);
            }

            // Save the resulting document.
            doc.Save(@"C:\Output\OleWithValidatedProgId.docx");
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
