using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Hello World! This is a sample document.");
        source.Save("input.docx", SaveFormat.Docx);

        // Load the document.
        Document doc = new Document("input.docx");

        // Perform find-and-replace: replace "World" with "Aspose".
        doc.Range.Replace("World", "Aspose", new FindReplaceOptions(FindReplaceDirection.Forward));

        // Save the modified document as Excel (XLSX).
        doc.Save("output.xlsx", SaveFormat.Xlsx);

        // Validate that the output file was created.
        if (!File.Exists("output.xlsx"))
        {
            throw new InvalidOperationException("Expected output XLSX was not created.");
        }
    }
}
