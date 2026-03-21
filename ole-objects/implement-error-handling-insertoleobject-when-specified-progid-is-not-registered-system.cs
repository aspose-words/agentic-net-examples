#nullable enable
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
            if (builder == null) throw new ArgumentNullException(nameof(builder));
            if (dataStream == null) throw new ArgumentNullException(nameof(dataStream));
            if (string.IsNullOrWhiteSpace(progId)) throw new ArgumentException("ProgId cannot be null or empty.", nameof(progId));

            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Console.WriteLine("OLE insertion is only supported on Windows platforms. Skipping insertion.");
                return null;
            }

            using (RegistryKey? key = Registry.ClassesRoot.OpenSubKey(progId))
            {
                if (key == null)
                {
                    Console.WriteLine($"ProgId \"{progId}\" is not registered on this machine. OLE object will not be inserted.");
                    return null;
                }
            }

            try
            {
                Shape oleShape = builder.InsertOleObject(dataStream, progId, asIcon, presentationStream);
                return oleShape;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to insert OLE object with ProgId \"{progId}\": {ex.Message}");
                return null;
            }
        }
    }

    public static class Example
    {
        public static void Run()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            const string excelFileName = "Sample.xlsx";

            // Ensure the sample file exists; if not, create a minimal placeholder.
            if (!File.Exists(excelFileName))
            {
                Console.WriteLine($"File \"{excelFileName}\" not found. Creating a placeholder file.");
                // Create a tiny empty Excel file (ZIP container with minimal structure) to avoid errors.
                // For simplicity, write a few bytes; Aspose.Words will treat it as a generic binary stream.
                File.WriteAllBytes(excelFileName, new byte[] { 0x50, 0x4B, 0x03, 0x04 }); // ZIP header
            }

            using (FileStream fileStream = File.Open(excelFileName, FileMode.Open, FileAccess.Read))
            {
                Shape? shape = OleHelper.InsertOleObjectWithCheck(
                    builder,
                    fileStream,
                    "Excel.Sheet.12",
                    asIcon: false,
                    presentationStream: null);

                if (shape != null)
                {
                    Console.WriteLine("OLE object inserted successfully.");
                }
                else
                {
                    Console.WriteLine("OLE object was not inserted.");
                }
            }

            doc.Save("Result.docx");
            Console.WriteLine("Document saved as Result.docx");
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            Example.Run();
        }
    }
}
