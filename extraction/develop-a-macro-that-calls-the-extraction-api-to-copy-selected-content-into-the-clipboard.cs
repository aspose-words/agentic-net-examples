using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    [STAThread]
    public static void Main()
    {
        // Create a sample source document with three paragraphs.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph to copy.");
        builder.Writeln("Third paragraph.");
        source.Save("source.docx");

        // Load the document we just created.
        Document loaded = new Document("source.docx");

        // Select the second paragraph (index 1) as the content to extract.
        Paragraph paragraphToCopy = loaded.FirstSection.Body.Paragraphs[1];
        if (paragraphToCopy == null)
            throw new InvalidOperationException("The paragraph to copy was not found.");

        // Extract the paragraph text.
        string extractedText = paragraphToCopy.GetText();

        // Instead of using System.Windows.Forms.Clipboard (which requires additional references),
        // write the extracted text to a deterministic file that represents the clipboard content.
        const string clipboardFile = "clipboard.txt";
        File.WriteAllText(clipboardFile, extractedText);

        // Verify that the file was created and contains the expected text.
        if (!File.Exists(clipboardFile))
            throw new InvalidOperationException("Failed to create the clipboard representation file.");

        string fileContent = File.ReadAllText(clipboardFile);
        if (fileContent != extractedText)
            throw new InvalidOperationException("The clipboard representation does not match the extracted content.");

        // Optional: output a confirmation message.
        Console.WriteLine("Paragraph extracted and written to '" + clipboardFile + "'.");
    }
}
