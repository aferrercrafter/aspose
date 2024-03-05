# Exercise 1

I have a Word document with multiple sections, and each section has different page
numbers. For example, Section 1 has page numbers 1 to 4, and Section 2 has page
numbers 6 to 9. How can I reset the page numbers for all sections in sequential order, I
tried using the property RestartPageNumbering, but I still don’t get the expected output?
(see attached Doc1.docx)


# Response

The ```RestartPageNumbering``` was indeed the correct path to follow, the problem you are facing is that the existing page numbers of the provided document are simple text (a paragraph node). Here is a example code snippet that will reset your page numbers: 

        ...
        var builder = new DocumentBuilder(document);

        var sectionIndex = 0;
        foreach (Section section in document.Sections)
        {
            section.PageSetup.RestartPageNumbering = false;
            builder.MoveToSection(1);

            var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            var footer2 = section.HeadersFooters[HeaderFooterType.FooterFirst];
            var footer3 = section.HeadersFooters[HeaderFooterType.FooterEven];
            var hasPageNumber = ContainsPageNumber(footer);
            if (!hasPageNumber)
            {
                builder.MoveToSection(sectionIndex);
                builder.MoveTo(footer.LastChild);
                builder.InsertField(FieldType.FieldPage, true);
            }

            sectionIndex++;
        }
        ...
        
        ...
        static bool ContainsPageNumber(HeaderFooter footer)
        {
            var fields = footer.GetChildNodes(NodeType.FieldStart, true);

            return fields.Any(x => ((FieldStart)x).FieldType == FieldType.FieldPage);
        }