using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Load the template document that contains multiple sections.
        // Some sections will be populated from JSON data, others are static.
        Document doc = new Document("Template.docx");

        // -----------------------------------------------------------------
        // 1. Perform mail merge for the sections that are generated from JSON.
        // -----------------------------------------------------------------
        // Assume the template has a mail‑merge region named "JsonData".
        // The JSON data is first deserialized into a DataTable (example shown).
        DataTable jsonTable = GetJsonDataAsDataTable();

        // Execute the mail merge only for the "JsonData" region.
        // This will create new sections (or fill existing ones) with the data.
        doc.MailMerge.ExecuteWithRegions(jsonTable);

        // -----------------------------------------------------------------
        // 2. Remove empty paragraphs **only** in the sections that originated
        //    from the JSON mail‑merge operation.
        // -----------------------------------------------------------------
        // For demonstration we identify those sections by a marker text
        // placed in the first paragraph of each JSON‑generated section:
        // e.g. "[JSON]" – this can be any unique string or a bookmark.
        const string jsonMarker = "[JSON]";

        foreach (Section section in doc.Sections)
        {
            // Skip sections that do not contain the marker.
            Paragraph firstPara = section.Body.FirstParagraph;
            if (firstPara == null || !firstPara.GetText().Contains(jsonMarker))
                continue;

            // Iterate over a copy of the paragraph collection because we will
            // be removing nodes while iterating.
            ParagraphCollection paragraphs = section.Body.Paragraphs;
            for (int i = paragraphs.Count - 1; i >= 0; i--)
            {
                Paragraph para = paragraphs[i];

                // Do not delete the mandatory empty paragraph that terminates a section.
                if (para.IsEndOfSection && !para.HasChildNodes)
                    continue;

                // Consider a paragraph empty if its visible text (after trimming) is zero length.
                string text = para.GetText().Replace("\a", string.Empty) // remove paragraph mark
                                         .Replace("\r", string.Empty) // remove carriage return
                                         .Trim();

                if (string.IsNullOrEmpty(text))
                {
                    // Remove the empty paragraph from the document.
                    para.Remove();
                }
            }
        }

        // -----------------------------------------------------------------
        // 3. Save the resulting document.
        // -----------------------------------------------------------------
        doc.Save("Result.docx");
    }

    // Helper method: converts JSON to a DataTable.
    // Replace the implementation with actual JSON deserialization as needed.
    static DataTable GetJsonDataAsDataTable()
    {
        // Example JSON:
        // [
        //   { "FirstName": "John", "LastName": "Doe" },
        //   { "FirstName": "Jane", "LastName": "Smith" }
        // ]
        // For simplicity we construct the DataTable manually.
        DataTable table = new DataTable("JsonData");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");

        table.Rows.Add(new object[] { "John", "Doe" });
        table.Rows.Add(new object[] { "Jane", "Smith" });

        return table;
    }
}
