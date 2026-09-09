using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Plain‑text source containing numbered items.
        string plainText = 
            "Shopping List:\n" +
            "1 Milk\n" +
            "2 Eggs\n" +
            "3 Bread\n\n" +
            "Tasks:\n" +
            "1) Finish report\n" +
            "2) Call client\n" +
            "3) Schedule meeting\n";

        // Load the text with options that recognise list numbering (including whitespace delimiters).
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = true
        };

        // The constructor overload loads from a stream using the specified options.
        Document doc = new Document(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(plainText)), loadOptions);

        // Update list labels so that Word list numbers are correctly stored.
        doc.UpdateListLabels();

        // (Optional) Count how many paragraphs were recognised as list items.
        int detectedListItems = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                detectedListItems++;
        }

        // Save the document with proper Word list structures.
        doc.Save("ConvertedLists.docx");
    }
}
