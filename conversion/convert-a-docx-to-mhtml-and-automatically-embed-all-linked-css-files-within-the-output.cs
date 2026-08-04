using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputPath = "sample.docx";
        const string outputPath = "sample.mht";

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");
        builder.Writeln("This document will be converted to MHTML with embedded CSS.");
        doc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Prepare save options for MHTML.
        //    - Use HtmlSaveOptions with SaveFormat.Mhtml.
        //    - Set CssStyleSheetType to Embedded so CSS is placed inside a <style> tag.
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            CssStyleSheetType = CssStyleSheetType.Embedded
        };

        // -----------------------------------------------------------------
        // 4. Save the document as MHTML.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath, saveOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the output file was created and contains data.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("MHTML output file was not created.");

        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
            throw new InvalidOperationException("MHTML output file is empty.");

        // Optional: indicate success.
        Console.WriteLine($"Document successfully converted to MHTML: {outputPath}");
    }
}
