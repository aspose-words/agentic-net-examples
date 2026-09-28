using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path to the Excel file that will be inserted as an OLE object.
        string excelPath = "sample.xlsx";

        // Ensure the Excel file exists; if not, create an empty placeholder file.
        if (!File.Exists(excelPath))
        {
            File.WriteAllBytes(excelPath, new byte[0]);
        }

        // List of Word documents to process.
        string[] wordFiles = new string[]
        {
            "Doc1.docx",
            "Doc2.docx",
            "Doc3.docx"
        };

        // Output directory for the modified documents.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        foreach (string wordFile in wordFiles)
        {
            // Skip if the source Word file does not exist.
            if (!File.Exists(wordFile))
                continue;

            // Load the Word document.
            Document doc = new Document(wordFile);

            // Create a DocumentBuilder for inserting content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move to the end of the document (or any desired location).
            builder.MoveToDocumentEnd();

            // Insert the Excel OLE object using streams (required by the API version).
            using (FileStream oleStream = File.OpenRead(excelPath))
            using (MemoryStream displayNameStream = new MemoryStream(Encoding.UTF8.GetBytes("Sample Excel")))
            {
                // Overload: InsertOleObject(Stream oleStream, string progId, bool isObjectIcon, Stream displayName)
                builder.InsertOleObject(oleStream, "Excel.Sheet", false, displayNameStream);
            }

            // Save the modified document.
            string outputPath = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(wordFile) + "_WithExcel.docx");
            doc.Save(outputPath);
        }
    }
}
