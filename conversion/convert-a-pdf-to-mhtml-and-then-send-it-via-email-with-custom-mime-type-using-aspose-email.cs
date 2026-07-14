using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words PDF to MHTML conversion.");

        // Step 2: Save the document as PDF.
        string pdfPath = "sample.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Step 3: Load the PDF and convert it to MHTML.
        Document pdfDoc = new Document(pdfPath);
        string mhtmlPath = "sample.mht";
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportCidUrlsForMhtmlResources = true
        };
        pdfDoc.Save(mhtmlPath, mhtmlOptions);
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML file was not created.");

        // Step 4: Create a simple MIME email with the MHTML attached using a custom MIME type.
        string emlPath = "email.eml";
        string boundary = "BOUNDARY_" + Guid.NewGuid().ToString("N");

        // Read the MHTML file and encode it in Base64.
        byte[] mhtmlBytes = File.ReadAllBytes(mhtmlPath);
        string mhtmlBase64 = Convert.ToBase64String(mhtmlBytes, Base64FormattingOptions.InsertLineBreaks);

        // Build the MIME message.
        StringBuilder mimeBuilder = new StringBuilder();
        mimeBuilder.AppendLine("From: sender@example.com");
        mimeBuilder.AppendLine("To: recipient@example.com");
        mimeBuilder.AppendLine("Subject: PDF converted to MHTML");
        mimeBuilder.AppendLine("MIME-Version: 1.0");
        mimeBuilder.AppendLine($"Content-Type: multipart/mixed; boundary=\"{boundary}\"");
        mimeBuilder.AppendLine();
        mimeBuilder.AppendLine($"--{boundary}");
        mimeBuilder.AppendLine("Content-Type: text/plain; charset=\"utf-8\"");
        mimeBuilder.AppendLine("Content-Transfer-Encoding: 7bit");
        mimeBuilder.AppendLine();
        mimeBuilder.AppendLine("Please find the MHTML attachment.");
        mimeBuilder.AppendLine();
        mimeBuilder.AppendLine($"--{boundary}");
        mimeBuilder.AppendLine("Content-Type: application/x-custom-mhtml; name=\"sample.mht\"");
        mimeBuilder.AppendLine("Content-Transfer-Encoding: base64");
        mimeBuilder.AppendLine("Content-Disposition: attachment; filename=\"sample.mht\"");
        mimeBuilder.AppendLine();
        mimeBuilder.AppendLine(mhtmlBase64);
        mimeBuilder.AppendLine();
        mimeBuilder.AppendLine($"--{boundary}--");
        mimeBuilder.AppendLine();

        // Save the .eml file.
        File.WriteAllText(emlPath, mimeBuilder.ToString(), Encoding.UTF8);
        if (!File.Exists(emlPath))
            throw new InvalidOperationException("EML file was not created.");

        // Indicate successful completion.
        Console.WriteLine("Conversion and email creation completed successfully.");
    }
}
