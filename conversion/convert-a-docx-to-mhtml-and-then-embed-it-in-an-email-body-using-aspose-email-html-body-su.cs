using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX document locally.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello, this is a sample DOCX content.");
        const string docxPath = "sample.docx";
        sourceDoc.Save(docxPath, SaveFormat.Docx);

        // Verify the DOCX file was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The sample DOCX file was not created.");

        // Step 2: Load the DOCX document.
        Document doc = new Document(docxPath);

        // Step 3: Convert the document to MHTML and capture the output in a memory stream.
        using MemoryStream mhtmlStream = new MemoryStream();
        doc.Save(mhtmlStream, SaveFormat.Mhtml);

        // Ensure the stream contains data.
        if (mhtmlStream.Length == 0)
            throw new InvalidOperationException("MHTML conversion produced no data.");

        // Reset the stream position and read the MHTML content as a string.
        mhtmlStream.Position = 0;
        string mhtmlContent;
        using (StreamReader reader = new StreamReader(mhtmlStream))
        {
            mhtmlContent = reader.ReadToEnd();
        }

        // Step 4: Build a simple .eml file with the MHTML content as the HTML body.
        const string emlPath = "output.eml";
        string emailHeaders =
            "From: sender@example.com\r\n" +
            "To: recipient@example.com\r\n" +
            "Subject: DOCX converted to MHTML\r\n" +
            "MIME-Version: 1.0\r\n" +
            "Content-Type: text/html; charset=utf-8\r\n" +
            "\r\n";

        File.WriteAllText(emlPath, emailHeaders + mhtmlContent);

        // Verify the email file was created.
        if (!File.Exists(emlPath))
            throw new InvalidOperationException("The email file was not created.");

        // Optional cleanup (commented out to keep files for inspection).
        // File.Delete(docxPath);
        // File.Delete(emlPath);
    }
}
