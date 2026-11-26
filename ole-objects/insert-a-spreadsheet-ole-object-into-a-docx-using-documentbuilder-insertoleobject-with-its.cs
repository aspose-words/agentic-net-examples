using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Simple CSV data that will be embedded as an Excel OLE object.
        string csvData = "Header1,Header2,Header3\n1,2,3\n4,5,6";

        // Convert the CSV string to a memory stream.
        using (MemoryStream oleStream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
        {
            // Insert the OLE object using the Excel ProgId.
            // The fourth parameter is an optional icon stream; we pass null.
            builder.InsertOleObject(oleStream, "Excel.Sheet", false, (Stream)null);
        }

        // Save the document to the output file.
        doc.Save("Output.docx");
    }
}
