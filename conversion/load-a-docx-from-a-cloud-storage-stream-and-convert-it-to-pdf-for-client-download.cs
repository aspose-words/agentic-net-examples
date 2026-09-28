using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX file locally.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOCX content from cloud storage.");
        source.Save("input.docx", SaveFormat.Docx);

        // Simulate loading the DOCX from a cloud storage stream.
        using (FileStream fileStream = new FileStream("input.docx", FileMode.Open, FileAccess.Read))
        using (MemoryStream cloudStream = new MemoryStream())
        {
            fileStream.CopyTo(cloudStream);
            cloudStream.Position = 0;

            Document doc = new Document(cloudStream);

            // Convert the document to PDF and write to a simulated client response stream.
            using (MemoryStream responseStream = new MemoryStream())
            {
                doc.Save(responseStream, SaveFormat.Pdf);

                if (responseStream.Length == 0)
                {
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");
                }

                // Optionally save the PDF to a file to verify the conversion.
                File.WriteAllBytes("output.pdf", responseStream.ToArray());

                if (!File.Exists("output.pdf"))
                {
                    throw new InvalidOperationException("Expected output PDF was not created.");
                }
            }
        }
    }
}
