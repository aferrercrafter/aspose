# Exercise 2

While I’m reading the cell content using cell.getText().trim(), I’m losing the list context. Is
there any library method which will return the cell content in the same list format. I need
to create a new section for each row of the table using the content of the Col1 as title and
appending the content of the rest of columns in the body of the corresponding section
without losing format.

# Response

```Aspose.Words.NodeImporter```. It's buiild exactly for the scenario of importing nodes:

And provides you 3 different importing modes, including, keep the source formatting.

        NodeImporter importer = new NodeImporter(document, newDoc, ImportFormatMode.KeepSourceFormatting);

You use the table of the existing document as the reference to let the ```NodeImporter``` get the nodes copy without loosing the format, and use that copy to create your new documents sections. Here is a example code snippet that will provide you the expected output: 

        var sectionIndex = 0;
        var lastSectionIndex = table.Rows.Count - 1;
        foreach (Row row in table.Rows)
        {
            var cellIndex = 0;
            foreach (Cell cell in row.Cells)
            {
                if (cellIndex == 0)
                {
                    var destNode = importer.ImportNode(cell.FirstParagraph, true);
                    newDoc.Sections[sectionIndex].Body.AppendChild(destNode);
                }
                else
                {
                    var nodes = cell.GetChildNodes(NodeType.Paragraph, true);
                    foreach (var node in nodes)
                    {
                        var destNode = importer.ImportNode(node, true);
                        newDoc.Sections[sectionIndex].Body.AppendChild(destNode);
                    }
                }
                cellIndex++;
            }

            if (sectionIndex < lastSectionIndex)
            {
                builder.MoveToSection(sectionIndex);
                builder.MoveTo(newDoc.Sections[sectionIndex].Body.LastChild);
                builder.InsertBreak(BreakType.SectionBreakNewPage);

                sectionIndex++;
            }
        }

