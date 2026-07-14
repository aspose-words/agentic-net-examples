using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX document.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Hello Aspose.Words! This document will be converted to MHTML and embedded in an email.");
        const string docxPath = "sample.docx";
        sampleDoc.Save(docxPath, SaveFormat.Docx);

        // Step 2: Load the DOCX document.
        Document loadedDoc = new Document(docxPath);

        // Step 3: Convert the document to MHTML and store it in a memory stream.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Use CID URLs for resources to improve compatibility.
                ExportCidUrlsForMhtmlResources = true
            };
            loadedDoc.Save(mhtmlStream, mhtmlOptions);

            if (mhtmlStream.Length == 0)
                throw new InvalidOperationException("MHTML conversion produced an empty stream.");

            // Reset the stream position before reading.
            mhtmlStream.Position = 0;

            // Step 4: Read the MHTML content as a string.
            string mhtmlContent;
            using (StreamReader reader = new StreamReader(mhtmlStream))
            {
                mhtmlContent = reader.ReadToEnd();
            }

            // Step 5: Create a simple MIME email with the MHTML as the HTML body.
            // Since Aspose.Email is not available, we construct the .eml file manually.
            string emlPath = "output.eml";
            string emailHeaders =
                "From: sender@example.com\r\n" +
                "To: recipient@example.com\r\n" +
                "Subject: Document embedded as MHTML\r\n" +
                "MIME-Version: 1.0\r\n" +
                "Content-Type: text/html; charset=utf-8\r\n" +
                "\r\n";

            File.WriteAllText(emlPath, emailHeaders + mhtmlContent);

            // Step 6: Validate that the email file was created.
            if (!File.Exists(emlPath))
                throw new InvalidOperationException("The email file was not created as expected.");
        }
    }
}
