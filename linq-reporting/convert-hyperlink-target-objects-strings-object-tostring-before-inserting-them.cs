using System;
using Aspose.Words;
using Aspose.Words.Fields;

class HyperlinkDemo
{
    static void Main()
    {
        // ---------- Create a new document ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Objects that represent hyperlink components.
        // These could be any type; we convert them to string using ToString().
        object addressObject = new Uri("https://www.aspose.com/");
        object subAddressObject = "Section1";               // could be a bookmark name
        object targetObject = "NewWindow";                  // target frame name

        // Convert objects to string before using them.
        string address = addressObject.ToString();
        string subAddress = subAddressObject.ToString();
        string target = targetObject.ToString();

        // Insert a hyperlink using the DocumentBuilder.InsertHyperlink method.
        // The method returns a Field object; we cast it to FieldHyperlink to set extra properties.
        Field field = builder.InsertHyperlink("Aspose Website", address, false);
        FieldHyperlink hyperlink = (FieldHyperlink)field;

        // Set additional hyperlink properties using the string values.
        hyperlink.SubAddress = subAddress;   // e.g., a bookmark within the document
        hyperlink.Target = target;           // frame name for the link
        hyperlink.ScreenTip = "Open " + address + " in a new window";

        // Save the document to disk.
        const string outputPath = "HyperlinkDemo.docx";
        doc.Save(outputPath);

        // ---------- Load the saved document ----------
        Document loadedDoc = new Document(outputPath);

        // Iterate through all fields and print hyperlink details.
        foreach (Field f in loadedDoc.Range.Fields)
        {
            if (f.Type == FieldType.FieldHyperlink)
            {
                FieldHyperlink hl = (FieldHyperlink)f;
                Console.WriteLine("Display Text : " + hl.Result);
                Console.WriteLine("Address      : " + hl.Address);
                Console.WriteLine("SubAddress   : " + hl.SubAddress);
                Console.WriteLine("Target       : " + hl.Target);
                Console.WriteLine("ScreenTip    : " + hl.ScreenTip);
                Console.WriteLine();
            }
        }
    }
}
