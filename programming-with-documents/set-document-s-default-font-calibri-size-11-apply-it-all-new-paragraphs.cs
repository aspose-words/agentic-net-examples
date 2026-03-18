using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the document-wide default font to Calibri, size 11.
        // This default is applied to any new runs or paragraphs added to the document.
        doc.Styles.DefaultFont.Name = "Calibri";
        doc.Styles.DefaultFont.Size = 11;

        // Add some paragraphs to demonstrate that they inherit the default font.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph uses the default Calibri 11.");
        builder.Writeln("Second paragraph also uses the default Calibri 11.");

        // Save the document.
        doc.Save("DefaultFont.docx");
    }
}
