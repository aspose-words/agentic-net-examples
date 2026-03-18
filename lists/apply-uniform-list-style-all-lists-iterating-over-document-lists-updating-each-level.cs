using System;
using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Uniform list level formatting.
        const string uniformFontName = "Arial";
        const float uniformFontSize = 12f;
        Color uniformColor = Color.Black;
        const bool uniformBold = false;

        // Iterate over every list in the document.
        foreach (List list in doc.Lists)
        {
            // Iterate over each level of the current list.
            foreach (ListLevel level in list.ListLevels)
            {
                // Apply the uniform formatting to the level.
                level.Font.Name = uniformFontName;
                level.Font.Size = uniformFontSize;
                level.Font.Color = uniformColor;
                level.Font.Bold = uniformBold;
            }
        }

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
