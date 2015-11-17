using System;
using System.Collections.Generic;
using System.Linq;
using Novacode;

namespace Templ
{
    public abstract class TemplModule<T> : TemplModule where T : TemplMatchPara, new()
    {
        public new Func<T, TemplBuilder, T> CustomHandler = (m, docBuilder) => m;

        public TemplModule(String name, TemplBuilder docBuilder, string[] prefixes)
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
        /// <param name="m"></param>
        public abstract T Handler(T m);

        /// <summary>
        /// Container-specific finder.
        /// Get all instances from the document.
        /// </summary>
        /// <param name="rxp"></param>
        public abstract IEnumerable<T> FindAll(TemplRegex rxp);
    }

    public abstract class TemplModule
    {
        public string Name;
        public TemplBuilder DocBuilder;
        public TemplRegex[] Regexes;
        public Func<object, TemplBuilder, object> CustomHandler;

        public TemplModule(string name, TemplBuilder docBuilder, string[] prefixes)
        {
            foreach (var prefix in prefixes.Where(s => s.Contains(":")))
            {
                throw new FormatException($"Templ: Module \"{name}\": prefix \"{prefix}\" cannot contain the split character ':'");
            }
            Regexes = prefixes.Select(pre => new TemplRegex(pre)).ToArray();
            DocBuilder = docBuilder;
            Name = name;
        }

        public abstract void Build();
    }

    /* ---------------------------------------------------------------------- */
    /* ---------------------------------------------------------------------- */
    /* ---------------------------------------------------------------------- */

        /// <summary>
        ///This is a utility module, instantiated  by others to handle a sub-scope of the document.
        ///Don't add it to the main Modules collection. If you use the default constructor it won't do anything.
        /// </summary>
    public class TemplSubcollectionModule : TemplModule<TemplMatchText>
    {
        private string Path = "";
        private IEnumerable<Paragraph> Paragraphs = new List<Paragraph>();

        public TemplSubcollectionModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public TemplSubcollectionModule(string prefix = "$") : base("Subcollection", null, new string[] { prefix }) { }

        public void BuildFromScope(IEnumerable<Paragraph> paragraphs, string path)
        {
            Path = path;
            Paragraphs = paragraphs;
            Build();
            Path = "";
            paragraphs = new List<Paragraph>();
        }

        public override TemplMatchText Handler(TemplMatchText m)
        {
            if (m is TemplMatchPicture)
            {
                return (m as TemplMatchPicture).SetDescription(ParentPath(m.Name));
            }
            return m.ToText(ParentPath(m.Name));
        }
        public override IEnumerable<TemplMatchText> FindAll(TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, Paragraphs).Concat(TemplMatchPicture.Find(rxp, Paragraphs).Cast<TemplMatchText>());
        }
        public string ParentPath(string payload)
        {
            var nameParts = payload.Split(':');
            if (nameParts.Length < 2)
            {
                throw new FormatException($"Templ: Subcollection placeholder has too few :-separated fields: \"{payload}\"");
            }
            nameParts[1] = $"{Path}{(nameParts[1].Length == 0 || Path.Length == 0?"":".")}{nameParts[1]}";
            return $"{{{nameParts.Aggregate((a, b) => $"{a}:{b}") }}}";
        }
    }

    public class TemplSectionModule : TemplModule<TemplMatchSection>
    {
        public TemplSectionModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchSection Handler(TemplMatchSection m)
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
                throw new FormatException($"Templ: Section {m.Name} has a malformed tag body (too many occurrences of :)");
            }
            // 2 parts: second part is a bool expression in the model for 'delete section'
            m.Expired = TemplModelEntry.Get(DocBuilder.Model, nameParts[1]).AsType<bool>();
            return m;
        }
        public override IEnumerable<TemplMatchSection> FindAll(TemplRegex rxp)
        {
            // "Take 1", a section should only have one naming tag.
            return DocBuilder.Doc.GetSections().SelectMany(sec => TemplMatchSection.Find(rxp, sec).Take(1));
        }
    }
    public class TemplTableModule : TemplModule<TemplMatchTable>
    {
        public TemplTableModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchTable Handler(TemplMatchTable m)
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
                throw new FormatException($"Templ: Table \"{m.Name}\" has a malformed tag body (too many occurrences of :)");
            }
            // 2 parts: second part is a bool expression in the model for 'delete table'
            m.Expired = TemplModelEntry.Get(DocBuilder.Model, nameParts[1]).AsType<bool>();
            return m;
        }
        public override IEnumerable<TemplMatchTable> FindAll(TemplRegex rxp)
        {
            // "Take 1", a table should only have one naming tag.
            return DocBuilder.Doc.Tables.SelectMany(t => TemplMatchTable.Find(rxp, t).Take(1));
        }
    }
    public class TemplRepeatingRowModule : TemplModule<TemplMatchRow>
    {
        public TemplRepeatingRowModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchRow Handler(TemplMatchRow m)
        {
            var e = TemplModelEntry.Get(DocBuilder.Model, m.Name);
            var idx = m.RowIndex;
            if (idx < 0)
            {
                throw new IndexOutOfRangeException($"Templ: Row match does not have an index in a table for match \"{m.Name}\"");
            }
            m.RemovePlaceholder();
            foreach (var key in e.ToStringKeys())
            {
                var r = m.Table.InsertRow(m.Row, ++idx);
                new TemplSubcollectionModule().BuildFromScope(r.Paragraphs, $"{m.Name}[{key}]");
            }
            m.Row.Remove();
            m.Removed = true;
            return m;
        }
        public override IEnumerable<TemplMatchRow> FindAll(TemplRegex rxp)
        {
            // Takes 1 per row. A row should only have one collection tag.
            return DocBuilder.Doc.Tables.SelectMany(
                t => Enumerable.Range(0, t.RowCount).SelectMany(
                    i => TemplMatchRow.Find(rxp, t, i).Take(1)
                )
            );
        }
    }
    public class TemplRepeatingCellModule : TemplModule<TemplMatchCell>
    {
        public TemplRepeatingCellModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchCell Handler(TemplMatchCell m)
        {
            var width = m.Table.Rows.First().Cells.Count;
            var keys = TemplModelEntry.Get(DocBuilder.Model, m.Name).ToStringKeys();
            var nrows = keys.Count() / width + 1;
            var rowIdx = m.RowIndex;
            var keyIdx = m.CellIndex;
            if (rowIdx < 0)
            {
                throw new IndexOutOfRangeException($"Templ: Cell match does not have a row index in a table for match \"{m.Name}\"");
            }
            if (keyIdx < 0)
            {
                throw new IndexOutOfRangeException($"Templ: Cell match does not have a cell index in a table for match \"{m.Name}\"");
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
                    new TemplSubcollectionModule().BuildFromScope(cell.Paragraphs, $"{m.Name}[{keys[kIdx]}]");
                }
                else
                {
                    CellClear(cell);
                }
            }
            m.Row.Remove();
            m.Removed = true;
            return m;
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
        public override IEnumerable<TemplMatchCell> FindAll(TemplRegex rxp)
        {
            // Takes 1 per row. A row should only have one repeating cell.
            return DocBuilder.Doc.Tables.SelectMany(
                t => Enumerable.Range(0, t.RowCount).SelectMany(
                    rr => Enumerable.Range(0, t.Rows[rr].Cells.Count).SelectMany(
                        cc => TemplMatchCell.Find(rxp, t, rr, cc).Take(1)
                    )
                )
            );
        }
    }
    public class TemplRepeatingTextModule : TemplModule<TemplMatchText>
    {
        public TemplRepeatingTextModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchText Handler(TemplMatchText m)
        {
            var e = TemplModelEntry.Get(DocBuilder.Model, m.Name);
            m.RemovePlaceholder();
            foreach (var key in e.ToStringKeys())
            {
                var p = m.Paragraph.InsertParagraphAfterSelf(m.Paragraph);
                new TemplSubcollectionModule().BuildFromScope(new Paragraph[] { p }, $"{m.Name}[{key}]");
            }
            m.Paragraph.Remove(false);
            m.Removed = true;
            return m;
        }
        public override IEnumerable<TemplMatchText> FindAll(TemplRegex rxp)
        {
            // Takes 1 per paragraph. A paragraph should only have one repeat entry.
            return DocBuilder.Doc.Paragraphs.SelectMany(
                p => TemplMatchText.Find(rxp, p).Take(1)
            );
        }
    }
    public class TemplPictureReplaceModule : TemplModule<TemplMatchPicture>
    {
        public TemplPictureReplaceModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchPicture Handler(TemplMatchPicture m)
        {
            if (m.Name.Split(':').Length > 1)
            {
                throw new FormatException($"Templ: Template Picture {m.Name} has a malformed tag body (too many occurrences of :)");
            }
            var e = TemplModelEntry.Get(DocBuilder.Model, m.Name);
            var w = m.Picture.Width;
            // Single picture: add text placeholder, expire the placeholder picture
            if (e.Value is TemplGraphic)
            {
                return m.ToText($"{{pic:{m.Name}:{w}}}");
            }
            // Multiple pictures: add repeating list placeholder, expire the placeholder picture
            if (e.Value is TemplGraphic[] || e.Value is ICollection<TemplGraphic>)
            {
                return m.ToText($"{{li:{m.Name}}}{{$:pic::{w}}}");
            }
            throw new InvalidCastException($"Templ: Failed to retrieve picture(s) from the model at path \"{e.Path}\"; its actual type is \"{e.Type}\"");
        }
        public override IEnumerable<TemplMatchPicture> FindAll(TemplRegex rxp)
        {
            return TemplMatchPicture.Find(rxp, DocBuilder.Doc.Paragraphs);
        }
    }
    public class TemplPicturePlaceholderModule : TemplModule<TemplMatchText>
    {
        public TemplPicturePlaceholderModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchText Handler(TemplMatchText m)
        {
            var nameParts = m.Name.Split(':');
            int w = -1;
            // 3+ parts: placeholder name is malformed
            if (nameParts.Length > 2)
            {
                throw new FormatException($"Templ: Picture {m.Name} has a malformed tag body (too many occurrences of :)");
            }
            // 2 parts: second part is an int for the width
            if (nameParts.Length == 2)
            {
                if (!int.TryParse(nameParts[1], out w))
                {
                    throw new FormatException($"Templ: Picture {m.Name} has a non-integer value for width ({nameParts[1]})");
                }
            }
            var e = TemplModelEntry.Get(DocBuilder.Model, nameParts[0]);
            //Try as array, as collection, as single
            if (e.Value is TemplGraphic)
            {
                return m.ToPicture((e.Value as TemplGraphic).Load(DocBuilder.Doc), w);
            }
            if (e.Value is TemplGraphic[])
            {
                return m.ToPictures((e.Value as TemplGraphic[])
                        .Select(g => g.Load(DocBuilder.Doc)).ToArray(), w);
            }
            if (e.Value is ICollection<TemplGraphic>)
            {
                return m.ToPictures((e.Value as ICollection<TemplGraphic>)
                        .Select(g => g.Load(DocBuilder.Doc)).ToList(), w);
            }
            throw new InvalidCastException($"Templ: Failed to retrieve picture(s) from the model at path \"{e.Path}\"; its actual type is \"{e.Type}\"");
        }
        public override IEnumerable<TemplMatchText> FindAll(TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, DocBuilder.Doc.Paragraphs);
        }
    }
    public class TemplTextModule : TemplModule<TemplMatchText>
    {
        public TemplTextModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchText Handler(TemplMatchText m)
        {
            m.ToText(TemplModelEntry.Get(DocBuilder.Model, m.Name).ToString());
            return m;
        }
        public override IEnumerable<TemplMatchText> FindAll(TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, DocBuilder.Doc.Paragraphs);
        }
    }
    public class TemplTOCModule : TemplModule<TemplMatchText>
    {
        public const TableOfContentsSwitches Switches =
            TableOfContentsSwitches.O | TableOfContentsSwitches.H | TableOfContentsSwitches.Z | TableOfContentsSwitches.U;

        public TemplTOCModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }

        /// <summary>
        /// Requires user to open in Word and click "Update table of contents". We cannot auto-populate it here.
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        public override TemplMatchText Handler(TemplMatchText m)
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
        public override IEnumerable<TemplMatchText> FindAll(TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, DocBuilder.Doc.Paragraphs);
        }
    }
    public class TemplCommentsModule : TemplModule<TemplMatchText>
    {
        public TemplCommentsModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchText Handler(TemplMatchText m)
        {
            m.Expired = true;
            return m;
        }
        public override IEnumerable<TemplMatchText> FindAll(TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, DocBuilder.Doc.Paragraphs);
        }
    }
}