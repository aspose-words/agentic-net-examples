using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Instantiate a blank Word document.
        Document doc = new Document();

        // Save the document as a DOTX (Word template) file.
        // The Save method determines the format from the SaveFormat enum.
        doc.Save("Template.dotx", SaveFormat.Dotx);
    }
}
