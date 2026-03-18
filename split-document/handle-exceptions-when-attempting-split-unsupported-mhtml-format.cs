using System;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitMhtmlExample
{
    static void Main()
    {
        // Create a simple document with a heading to demonstrate splitting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
        builder.Writeln("Sample Heading");
        builder.ParagraphFormat.Style = doc.Styles["Normal"];
        builder.Writeln("Some body text.");

        // Configure HtmlSaveOptions for MHTML format and attempt to split the document.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Splitting is not supported for MHTML; this setting will trigger an exception.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph
        };

        try
        {
            // Attempt to save; Aspose.Words should throw UnsupportedFileFormatException.
            doc.Save("Output.mhtml", saveOptions);
            Console.WriteLine("Document saved successfully (unexpected).");
        }
        catch (UnsupportedFileFormatException ex)
        {
            // Handle the specific exception that indicates splitting is not allowed for MHTML.
            Console.WriteLine("Cannot split when saving to MHTML format:");
            Console.WriteLine(ex.Message);
        }
        catch (Exception ex)
        {
            // Fallback for any other unexpected errors.
            Console.WriteLine("An unexpected error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
