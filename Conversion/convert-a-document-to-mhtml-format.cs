using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source document (any format supported by Aspose.Words).
        Document doc = new Document("input.docx");

        // Save the document in MHTML (Web archive) format.
        // The SaveFormat enumeration value Mhtml specifies the target format.
        doc.Save("output.mht", SaveFormat.Mhtml);
    }
}
