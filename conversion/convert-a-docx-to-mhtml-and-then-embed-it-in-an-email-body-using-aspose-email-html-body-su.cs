using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello, this is a sample DOCX content.");
        sourceDoc.Save("sample.docx", SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document("sample.docx");

        // -----------------------------------------------------------------
        // 3. Convert the DOCX to MHTML and read the content.
        // -----------------------------------------------------------------
        string mhtmlContent;
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            doc.Save(mhtmlStream, SaveFormat.Mhtml);

            if (mhtmlStream.Length == 0)
                throw new InvalidOperationException("MHTML conversion produced no data.");

            mhtmlStream.Position = 0;
            using (StreamReader reader = new StreamReader(mhtmlStream))
            {
                mhtmlContent = reader.ReadToEnd();
            }
        }

        // -----------------------------------------------------------------
        // 4. Build a simple RFC‑822 email message with the MHTML as the HTML body.
        // -----------------------------------------------------------------
        string emailPath = "email.eml";
        using (StreamWriter writer = new StreamWriter(emailPath, false))
        {
            writer.WriteLine("From: sender@example.com");
            writer.WriteLine("To: recipient@example.com");
            writer.WriteLine("Subject: Test Email with MHTML Body");
            writer.WriteLine("MIME-Version: 1.0");
            writer.WriteLine("Content-Type: text/html; charset=utf-8");
            writer.WriteLine(); // Blank line separates headers from body.
            writer.WriteLine(mhtmlContent);
        }

        // -----------------------------------------------------------------
        // 5. Verify that the email file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(emailPath))
            throw new InvalidOperationException("The email file was not created.");
    }
}
