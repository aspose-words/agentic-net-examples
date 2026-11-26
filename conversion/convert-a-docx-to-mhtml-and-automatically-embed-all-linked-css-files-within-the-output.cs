using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX document locally.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document for DOCX to MHTML conversion.");
        sourceDoc.Save("input.docx", SaveFormat.Docx);

        // Step 2: Load the DOCX document that was just created.
        Document doc = new Document("input.docx");

        // Step 3: Save the document as MHTML.
        // Aspose.Words automatically embeds linked CSS when saving to MHTML.
        doc.Save("output.mhtml", SaveFormat.Mhtml);

        // Step 4: Validate that the MHTML file was created and is not empty.
        if (!File.Exists("output.mhtml"))
        {
            throw new InvalidOperationException("Expected output MHTML file was not created.");
        }

        FileInfo fileInfo = new FileInfo("output.mhtml");
        if (fileInfo.Length == 0)
        {
            throw new InvalidOperationException("The generated MHTML file is empty.");
        }

        // Indicate successful completion.
        Console.WriteLine("DOCX successfully converted to MHTML with embedded CSS.");
    }
}
