using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document that will be converted to MHTML.");
        sourceDoc.Save("input.docx", SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Save the document as MHTML. The default behavior embeds linked CSS files.
        string outputPath = "output.mhtml";
        doc.Save(outputPath, SaveFormat.Mhtml);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The MHTML output file was not created.");
        }

        // Verify that the file is not empty.
        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
        {
            throw new InvalidOperationException("The MHTML output file is empty.");
        }

        // Indicate successful completion.
        Console.WriteLine("Conversion completed successfully.");
    }
}
