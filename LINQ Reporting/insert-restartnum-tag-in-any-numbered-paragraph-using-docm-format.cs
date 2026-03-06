using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new document or load an existing one
        Document doc = new Document(); // or new Document("Input.docx");

        // Use DocumentBuilder to add some numbered paragraphs for demonstration
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.ApplyNumberDefault(); // start a numbered list
        builder.Writeln("First numbered paragraph");
        builder.Writeln("Second numbered paragraph");
        builder.Writeln("Third numbered paragraph");

        // Iterate through all paragraphs in the document
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Check if the paragraph is part of a list (i.e., it is numbered)
            if (para.IsListItem)
            {
                // The List object associated with the paragraph contains the
                // IsRestartAtEachSection property. Setting it to true adds the
                // <w:restartNum/> tag to the underlying DOCM XML for that list.
                List list = para.ListFormat.List;
                if (list != null)
                {
                    list.IsRestartAtEachSection = true;
                }
            }
        }

        // Save the document in DOCM format (macro‑enabled Word document)
        doc.Save("Output.docm");
    }
}
