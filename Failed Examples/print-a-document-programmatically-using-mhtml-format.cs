// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build simple content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Save the document in MHTML format.
        // The Save(string, SaveFormat) rule is used here.
        string mhtmlFile = "HelloWorld.mhtml";
        doc.Save(mhtmlFile, SaveFormat.Mhtml);

        // Print the document to the default printer.
        // The Print() method follows the provided Print rule.
        doc.Print();

        Console.WriteLine("Document saved as MHTML and sent to printer.");
    }
}
