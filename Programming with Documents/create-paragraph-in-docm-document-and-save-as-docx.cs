using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document (initially in the default format).
        Document doc = new Document();

        // Use DocumentBuilder to insert a paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a paragraph added to a DOCM document.");

        // Save the document as DOCX. The file extension determines the save format.
        doc.Save("Result.docx");
    }
}
