// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing.Printing;

class PrintDotmExample
{
    static void Main()
    {
        // 1. Create a new blank document (uses the Document() constructor rule).
        Document doc = new Document();

        // 2. Add some content using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello from a DOTM template!");

        // 3. Save the document as a macro‑enabled template (.dotm).
        //    The Save(string) method determines the format from the file extension.
        string templatePath = "Template.dotm";
        doc.Save(templatePath); // uses the Save(string) rule.

        // 4. Load the saved .dotm file back into a Document object.
        Document loadedDoc = new Document(templatePath); // uses the Document(string) constructor rule.

        // 5. Print the loaded document to the default printer.
        //    The Print() method prints the whole document without UI.
        loadedDoc.Print(); // uses the Print() rule.
    }
}
