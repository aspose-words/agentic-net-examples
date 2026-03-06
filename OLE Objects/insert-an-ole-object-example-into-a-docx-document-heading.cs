using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a heading for the OLE object section.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Embedding an OLE Object");

        // Return to normal paragraph style for the following content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Path to the file that will be embedded as an OLE object.
        // Replace with an existing file on your system (e.g., an Excel workbook).
        string oleFilePath = @"C:\MyData\Sample.xlsx";

        // Insert the OLE object:
        // - isLinked = false  -> embed the file.
        // - asIcon   = true   -> display the object as an icon.
        // - presentation = null -> use the default icon provided by Aspose.Words.
        builder.InsertOleObject(oleFilePath, false, true, null);

        // Save the resulting document.
        doc.Save(@"C:\Output\OleObjectExample.docx");
    }
}
