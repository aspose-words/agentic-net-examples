using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

class ElementAtExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a collection with at least five items.
        List<string> items = new List<string>
        {
            "First",
            "Second",
            "Third",
            "Fourth",
            "Fifth",
            "Sixth"
        };

        // Retrieve the fifth item (zero‑based index 4) using LINQ's ElementAt method.
        string fifthItem = items.ElementAt(4);

        // Insert the retrieved value into the document.
        builder.Writeln($"The fifth item in the collection is: {fifthItem}");

        // Save the document to disk.
        doc.Save("ElementAtExample.docx");
    }
}
