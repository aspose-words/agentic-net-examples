using System;
using System.IO;
using System.Runtime.InteropServices;
using Aspose.Words;
using Aspose.Words.Drawing;
using Microsoft.Win32;

namespace OleObjectInsertion
{
    public static class OleHelper
    {
        /// <summary>
        /// Inserts an OLE object into the document only if the specified ProgId is registered on the system.
        /// Returns the created Shape or null when the ProgId is not found or the platform does not support OLE.
        /// </summary>
        public static Shape? InsertOleObjectWithCheck(DocumentBuilder builder,
                                                      Stream dataStream,
                                                      string progId,
                                                      bool asIcon,
                                                      Stream? presentationStream = null)
        {
            // Validate arguments.
            if (builder == null) throw new ArgumentNullException(nameof(builder));
            if (dataStream == null) throw new ArgumentNullException(nameof(dataStream));
            if (string.IsNullOrWhiteSpace(progId)) throw new ArgumentException("ProgId cannot be null or empty.", nameof(progId));

            // OLE registration is a Windows‑only feature.
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Console.WriteLine("OLE insertion is only supported on Windows platforms. Skipping insertion.");
                return null;
            }

            // Check whether the ProgId exists in the registry (HKEY_CLASSES_ROOT\<ProgId>).
            using (RegistryKey? key = Registry.ClassesRoot.OpenSubKey(progId))
            {
                if (key == null)
                {
                    // ProgId not registered – handle gracefully.
                    Console.WriteLine($"ProgId \"{progId}\" is not registered on this machine. OLE object will not be inserted.");
                    return null;
                }
            }

            // ProgId is present – attempt insertion.
            try
            {
                // InsertOleObject is an existing Aspose.Words method; we use it directly.
                Shape oleShape = builder.InsertOleObject(dataStream, progId, asIcon, presentationStream);
                return oleShape;
            }
            catch (Exception ex)
            {
                // Unexpected errors (e.g., corrupted stream) are caught here.
                Console.WriteLine($"Failed to insert OLE object with ProgId \"{progId}\": {ex.Message}");
                return null;
            }
        }
    }

    public static class Example
    {
        public static void Run()
        {
            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Prepare a stream containing the data to embed (e.g., an Excel file).
            using (FileStream fileStream = File.Open("Sample.xlsx", FileMode.Open, FileAccess.Read))
            {
                // Attempt to insert the OLE object with error handling.
                Shape? shape = OleHelper.InsertOleObjectWithCheck(
                    builder,
                    fileStream,
                    "Excel.Sheet.12",   // ProgId for Excel worksheets.
                    asIcon: false,
                    presentationStream: null);

                if (shape != null)
                {
                    Console.WriteLine("OLE object inserted successfully.");
                }
            }

            // Save the document.
            doc.Save("Result.docx");
        }
    }

    internal class Program
    {
        // Entry point required by the compiler.
        private static void Main(string[] args)
        {
            Example.Run();
        }
    }
}
