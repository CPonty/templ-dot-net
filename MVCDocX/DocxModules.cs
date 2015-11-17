using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Novacode;

namespace DocXMVC
{
    public abstract class DocxModule<T> : DocxModule where T : DocxMatchPara, new()
    {
        public new Func<T, DocxBuilder, T> CustomHandler = (m, docBuilder) => m;

        public DocxModule(String name, DocxBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }

        public void BuildFromScope(IEnumerable<T> scope)
        {
            // Note how we are constantly "committing" the changes using ToList().
            // This ensures order is preserved (e.g. all "finds" happen before all "handler"s)
            scope.Select( m => Handler(m)).ToList()
                 .Where(  m =>!m.RemoveExpired()).ToList()
                 .Select( m => CustomHandler(m, DocBuilder)).ToList()
                 .ForEach(m => m.RemoveExpired());
        }
        public override void Build()
        {
            BuildFromScope(Regexes.SelectMany(rxp => FindAll(rxp)));
        }
        /// <summary>
        /// Container-specific handler.
        /// Modify the underlying container; mark expired to delete
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="m"></param>
        /// <param name="docBuilder"></param>
        public abstract T Handler(T m);

        /// <summary>
        /// Container-specific finder.
        /// Get all instances from the document.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rxp"></param>
        public abstract IEnumerable<T> FindAll(DocxRegex rxp);
    }

    public abstract class DocxModule
    {
        public string Name;
        public DocxBuilder DocBuilder;
        public DocxRegex[] Regexes;
        public Func<object, DocxBuilder, object> CustomHandler;

        public DocxModule(string name, DocxBuilder docBuilder, string[] prefixes)
        {
            foreach (var prefix in prefixes.Where(s => s.Contains(":")))
            {
                throw new FormatException($"Docx: Module \"{name}\": prefix \"{prefix}\" cannot contain the split character ':'");
            }
            Regexes = prefixes.Select(pre => new DocxRegex(pre)).ToArray();
            DocBuilder = docBuilder;
            Name = name;
        }

        public abstract void Build();
    }

    /* ---------------------------------------------------------------------- */
    /* ---------------------------------------------------------------------- */
    /* ---------------------------------------------------------------------- */

    /*  This is a utility module, instantiated  by others to handle a sub-scope of the document.
        Don't add it to the main Modules collection. If you use the default constructor it won't do anything. */
    public class DocxSubcollectionModule : DocxModule<DocxMatchText>
    {
        private string Path = "";
        private IEnumerable<Paragraph> Paragraphs = new List<Paragraph>();

        public DocxSubcollectionModule(string name, DocxBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public DocxSubcollectionModule(string prefix = "$") : base("Subcollection", null, new string[] { prefix }) { }

        public void BuildFromScope(IEnumerable<Paragraph> paragraphs, string path)
        {
            Path = path;
            Paragraphs = paragraphs;
            Build();
            Path = "";
            paragraphs = new List<Paragraph>();
        }

        public override DocxMatchText Handler(DocxMatchText m)
        {
            if (m is DocxMatchPicture)
            {
                return (m as DocxMatchPicture).SetDescription(ParentPath(m.Name));
            }
            return m.ToText(ParentPath(m.Name));
        }
        public override IEnumerable<DocxMatchText> FindAll(DocxRegex rxp)
        {
            return DocxMatchText.Find(rxp, Paragraphs).Concat(DocxMatchPicture.Find(rxp, Paragraphs).Cast<DocxMatchText>());
        }
        public string ParentPath(string payload)
        {
            var nameParts = payload.Split(':');
            if (nameParts.Length < 2)
            {
                throw new FormatException($"Docx: Subcollection placeholder has too few :-separated fields: \"{payload}\"");
            }
            nameParts[1] = $"{Path}{(nameParts[1].Length == 0 || Path.Length == 0?"":".")}{nameParts[1]}";
            //TODO remove '|' once it works
            return $"{{{nameParts.Aggregate((a, b) => $"{a}:{b}") }}}";
        }
    }

    public class DocxSectionModule : DocxModule<DocxMatchSection>
    {
        public DocxSectionModule(string name, DocxBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override DocxMatchSection Handler(DocxMatchSection m)
        {
            m.RemovePlaceholder();
            var nameParts = m.Name.Split(':');
            // 1 part: placeholder name is just a label, nothing to do.
            // Users can implement their own section handler to switch off this label.
            if (nameParts.Length == 1)
            {
                return m;
            }
            // 3+ parts: placeholder name is malformed
            if (nameParts.Length > 2)
            {
                throw new FormatException($"Docx: Section {m.Name} has a malformed tag body (too many occurrences of :)");
            }
            // 2 parts: second part is a bool expression in the model for 'delete section'
            m.Expired = DocxModelEntry.Get(DocBuilder.Model, nameParts[1]).AsType<bool>();
            return m;
        }
        public override IEnumerable<DocxMatchSection> FindAll(DocxRegex rxp)
        {
            // "Take 1", a section should only have one naming tag.
            return DocBuilder.Doc.GetSections().SelectMany(sec => DocxMatchSection.Find(rxp, sec).Take(1));
        }
    }
    public class DocxTableModule : DocxModule<DocxMatchTable>
    {
        public DocxTableModule(string name, DocxBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override DocxMatchTable Handler(DocxMatchTable m)
        {
            m.RemovePlaceholder();
            var nameParts = m.Name.Split(':');
            // 1 part: placeholder name is just a label, nothing to do.
            // Users can implement their own table handler to switch off this label.
            if (nameParts.Length == 1)
            {
                return m;
            }
            // 3+ parts: placeholder name is malformed
            if (nameParts.Length > 2)
            {
                throw new FormatException($"Docx: Table \"{m.Name}\" has a malformed tag body (too many occurrences of :)");
            }
            // 2 parts: second part is a bool expression in the model for 'delete table'
            m.Expired = DocxModelEntry.Get(DocBuilder.Model, nameParts[1]).AsType<bool>();
            return m;
        }
        public override IEnumerable<DocxMatchTable> FindAll(DocxRegex rxp)
        {
            // "Take 1", a table should only have one naming tag.
            return DocBuilder.Doc.Tables.SelectMany(t => DocxMatchTable.Find(rxp, t).Take(1));
        }
    }
    public class DocxRepeatingRowModule : DocxModule<DocxMatchRow>
    {
        public DocxRepeatingRowModule(string name, DocxBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override DocxMatchRow Handler(DocxMatchRow m)
        {
            var e = DocxModelEntry.Get(DocBuilder.Model, m.Name);
            var idx = m.RowIndex;
            if (idx < 0)
            {
                throw new IndexOutOfRangeException($"Docx: Row match does not have an index in a table for match \"{m.Name}\"");
            }
            m.RemovePlaceholder();
            foreach (var key in e.ToStringKeys())
            {
                var r = m.Table.InsertRow(m.Row, ++idx);
                new DocxSubcollectionModule().BuildFromScope(r.Paragraphs, $"{m.Name}[{key}]");
            }
            m.Row.Remove();
            m.Removed = true;
            return m;
        }
        public override IEnumerable<DocxMatchRow> FindAll(DocxRegex rxp)
        {
            // Takes 1 per row. A row should only have one collection tag.
            return DocBuilder.Doc.Tables.SelectMany(
                t => Enumerable.Range(0, t.RowCount).SelectMany(
                    i => DocxMatchRow.Find(rxp, t, i).Take(1)
                )
            );
            //return DocxMatchRow.Find(rxp, DocBuilder.Doc.Tables);
        }
    }
    public class DocxRepeatingCellModule : DocxModule<DocxMatchCell>
    {
        public DocxRepeatingCellModule(string name, DocxBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override DocxMatchCell Handler(DocxMatchCell m)
        {
            var width = m.Table.Rows.First().Cells.Count;
            var keys = DocxModelEntry.Get(DocBuilder.Model, m.Name).ToStringKeys();
            var nrows = keys.Count() / width + 1;
            var rowIdx = m.RowIndex;
            var keyIdx = m.CellIndex;
            if (rowIdx < 0)
            {
                throw new IndexOutOfRangeException($"Docx: Cell match does not have a row index in a table for match \"{m.Name}\"");
            }
            if (keyIdx < 0)
            {
                throw new IndexOutOfRangeException($"Docx: Cell match does not have a cell index in a table for match \"{m.Name}\"");
            }
            m.RemovePlaceholder();
            for (int n=0; n< width; n++)
            {
                if (n != keyIdx)
                {
                    CellCopy(m.Cell, m.Row.Cells[n]);
                }
            }
            Row row = m.Row;
            for (int kIdx = 0; kIdx < nrows*width; kIdx++)
            {
                if (kIdx % width == 0)
                {
                    row = m.Table.InsertRow(m.Row, ++rowIdx);
                }
                Cell cell = row.Cells[kIdx % width];
                if (kIdx < keys.Count())
                {
                    new DocxSubcollectionModule().BuildFromScope(cell.Paragraphs, $"{m.Name}[{keys[kIdx]}]");
                }
                else
                {
                    CellClear(cell);
                }
            }
            m.Row.Remove();
            m.Removed = true;
            return m;
            //TODO
        }
        /// <summary>
        /// Clear text and images from all paragraphs in cell.
        ///
        /// Optional: Delete all paragraph objects.
        /// Keep in mind, cells with zero paragraphs are considered malformed by Word!
        /// Also keep in mind, you will lose formatting info (e.g. font).
        ///
        /// Special cases of content other than text or images may not be removed;
        /// If this becomes a problem, we can develop it.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="deleteAllParagraphs"></param>
        public static void CellClear(Cell cell, bool deleteAllParagraphs = false)
        {
            // Remove all but first paragraph
            cell.Paragraphs.Skip(deleteAllParagraphs ? 0 : 1).ToList().ForEach(p => cell.RemoveParagraph(p));
            if (!deleteAllParagraphs)
            {
                var p = cell.Paragraphs.Last();
                if (p.Text.Length > 0)
                {
                    p.RemoveText(0, p.Text.Length);
                }
            }
            cell.Pictures.ForEach(pic => pic.Remove());
        }

        /// <summary>
        /// Copy (text) contents of srcCell into dstCell.
        /// Clears text of dstCell plus all but first paragraph instance before copying.
        /// </summary>
        /// <param name="srcCell"></param>
        /// <param name="dstCell"></param>
        public static void CellCopy(Cell srcCell, Cell dstCell)
        {
            CellClear(dstCell, true);
            srcCell.Paragraphs.ToList().ForEach(p => dstCell.InsertParagraph(p));
        }
        public override IEnumerable<DocxMatchCell> FindAll(DocxRegex rxp)
        {
            // Takes 1 per row. A row should only have one repeating cell.
            return DocBuilder.Doc.Tables.SelectMany(
                t => Enumerable.Range(0, t.RowCount).SelectMany(
                    rr => Enumerable.Range(0, t.Rows[rr].Cells.Count).SelectMany(
                        cc => DocxMatchCell.Find(rxp, t, rr, cc).Take(1)
                    )
                )
            );
            //return DocxMatchCell.Find(rxp, DocBuilder.Doc.Tables);
        }
    }
    public class DocxRepeatingTextModule : DocxModule<DocxMatchText>
    {
        public DocxRepeatingTextModule(string name, DocxBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override DocxMatchText Handler(DocxMatchText m)
        {
            var e = DocxModelEntry.Get(DocBuilder.Model, m.Name);
            m.RemovePlaceholder();
            foreach (var key in e.ToStringKeys())
            {
                var p = m.Paragraph.InsertParagraphAfterSelf(m.Paragraph);
                new DocxSubcollectionModule().BuildFromScope(new Paragraph[] { p }, $"{m.Name}[{key}]");
            }
            m.Paragraph.Remove(false);
            m.Removed = true;
            return m;
        }
        public override IEnumerable<DocxMatchText> FindAll(DocxRegex rxp)
        {
            // Takes 1 per paragraph. A paragraph should only have one repeat entry.
            return DocBuilder.Doc.Paragraphs.SelectMany(
                p => DocxMatchText.Find(rxp, p).Take(1)
            );
        }
    }
    public class DocxPictureReplaceModule : DocxModule<DocxMatchPicture>
    {
        public DocxPictureReplaceModule(string name, DocxBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override DocxMatchPicture Handler(DocxMatchPicture m)
        {
            if (m.Name.Split(':').Length > 1)
            {
                throw new FormatException($"Docx: Template Picture {m.Name} has a malformed tag body (too many occurrences of :)");
            }
            var e = DocxModelEntry.Get(DocBuilder.Model, m.Name);
            var w = m.Picture.Width;
            // Single picture: add text placeholder, expire the placeholder picture
            if (e.Value is DocxGraphic)
            {
                return m.ToText($"{{pic:{m.Name}:{w}}}");
            }
            // Multiple pictures: add repeating list placeholder, expire the placeholder picture
            if (e.Value is DocxGraphic[] || e.Value is ICollection<DocxGraphic>)
            {
                return m.ToText($"{{li:{m.Name}}}{{$:pic::{w}}}");
            }
            throw new InvalidCastException($"Docx: Failed to retrieve picture(s) from the model at path \"{e.Path}\"; its actual type is \"{e.Type}\"");
        }
        public override IEnumerable<DocxMatchPicture> FindAll(DocxRegex rxp)
        {
            return DocxMatchPicture.Find(rxp, DocBuilder.Doc.Paragraphs);
        }
    }
    public class DocxPicturePlaceholderModule : DocxModule<DocxMatchText>
    {
        public DocxPicturePlaceholderModule(string name, DocxBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override DocxMatchText Handler(DocxMatchText m)
        {
            var nameParts = m.Name.Split(':');
            int w = -1;
            // 3+ parts: placeholder name is malformed
            if (nameParts.Length > 2)
            {
                throw new FormatException($"Docx: Picture {m.Name} has a malformed tag body (too many occurrences of :)");
            }
            // 2 parts: second part is an int for the width
            if (nameParts.Length == 2)
            {
                if (!int.TryParse(nameParts[1], out w))
                {
                    throw new FormatException($"Docx: Picture {m.Name} has a non-integer value for width ({nameParts[1]})");
                }
            }
            var e = DocxModelEntry.Get(DocBuilder.Model, nameParts[0]);
            //Try as array, as collection, as single
            if (e.Value is DocxGraphic)
            {
                return m.ToPicture((e.Value as DocxGraphic).Load(DocBuilder.Doc), w);
            }
            if (e.Value is DocxGraphic[])
            {
                return m.ToPictures((e.Value as DocxGraphic[])
                        .Select(g => g.Load(DocBuilder.Doc)).ToArray(), w);
            }
            if (e.Value is ICollection<DocxGraphic>)
            {
                return m.ToPictures((e.Value as ICollection<DocxGraphic>)
                        .Select(g => g.Load(DocBuilder.Doc)).ToList(), w);
            }
            throw new InvalidCastException($"Docx: Failed to retrieve picture(s) from the model at path \"{e.Path}\"; its actual type is \"{e.Type}\"");
        }
        public override IEnumerable<DocxMatchText> FindAll(DocxRegex rxp)
        {
            return DocxMatchText.Find(rxp, DocBuilder.Doc.Paragraphs);
        }
    }
    public class DocxTextModule : DocxModule<DocxMatchText>
    {
        public DocxTextModule(string name, DocxBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override DocxMatchText Handler(DocxMatchText m)
        {
            m.ToText(DocxModelEntry.Get(DocBuilder.Model, m.Name).ToString());
            return m;
        }
        public override IEnumerable<DocxMatchText> FindAll(DocxRegex rxp)
        {
            return DocxMatchText.Find(rxp, DocBuilder.Doc.Paragraphs);
        }
    }
    public class DocxTOCModule : DocxModule<DocxMatchText>
    {
        public const TableOfContentsSwitches Switches =
            TableOfContentsSwitches.O | TableOfContentsSwitches.H | TableOfContentsSwitches.Z | TableOfContentsSwitches.U;

        public DocxTOCModule(string name, DocxBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }

        /// <summary>
        /// Requires user to open in Word and click "Update table of contents". We cannot auto-populate it here.
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public override DocxMatchText Handler(DocxMatchText m)
        {
            if (m.Removed)
            {
                return m;
            }
            DocBuilder.Doc.InsertTableOfContents(m.Paragraph, m.Name, Switches);
            /* Additional options:
                string headerStyle = null
                int maxIncludeLevel = 3
                int? rightTabPos = null) */
            m.RemovePlaceholder();
            return m;
        }
        public override IEnumerable<DocxMatchText> FindAll(DocxRegex rxp)
        {
            return DocxMatchText.Find(rxp, DocBuilder.Doc.Paragraphs);
        }
    }
    public class DocxCommentsModule : DocxModule<DocxMatchText>
    {
        public DocxCommentsModule(string name, DocxBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override DocxMatchText Handler(DocxMatchText m)
        {
            m.Expired = true;
            return m;
        }
        public override IEnumerable<DocxMatchText> FindAll(DocxRegex rxp)
        {
            return DocxMatchText.Find(rxp, DocBuilder.Doc.Paragraphs);
        }
    }
}