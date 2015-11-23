using System;
using System.Collections.Generic;
using System.Linq;
using Novacode;

namespace TemplNET
{
        /// <summary>
        ///This is a utility module, instantiated  by others to handle a sub-scope of the document.
        ///Don't add it to the main Modules collection. If you use the default constructor it won't do anything.
        /// </summary>
    public class TemplSubcollectionModule : TemplModule<TemplMatchText>
    {
        private string Path = "";
        private IEnumerable<Paragraph> Paragraphs = new List<Paragraph>();

        public TemplSubcollectionModule(string name = "Subcollection", string prefix = "$")
            : base(name, prefix)
        {
            MinFields = 2;
            MaxFields = 99;
        }

        public void BuildFromScope(DocX doc, object model, IEnumerable<Paragraph> paragraphs, string path)
        {
            Path = path;
            Paragraphs = paragraphs;
            Build(doc, model);
            Path = "";
            paragraphs = new List<Paragraph>();
        }

        public override TemplMatchText Handler(DocX doc, object model, TemplMatchText m)
        {
            if (m is TemplMatchPicture)
            {
                return (m as TemplMatchPicture).SetDescription(ParentPath(m));
            }
            return m.ToText(ParentPath(m));
        }
        public override IEnumerable<TemplMatchText> FindAll(DocX doc, TemplRegex rxp)
        {
            // Get both text and picture matches. Contact is possible because MatchText is MatchPicture's base type.
            return TemplMatchText.Find(rxp, Paragraphs).Concat(
                   TemplMatchPicture.Find(rxp, Paragraphs).Cast<TemplMatchText>());
        }
        public string ParentPath(TemplMatchText m)
        {
            var fields = m.Fields;
            fields[1] = $"{Path}{(fields[1].Length == 0 || Path.Length == 0?"":".")}{fields[1]}";
            return $"{TemplConst.MatchOpen}{fields.Aggregate((a, b) => $"{a}:{b}") }{TemplConst.MatchClose}";
        }
    }

    public class TemplSectionModule : TemplModule<TemplMatchSection>
    {
        public TemplSectionModule(string name, string prefix)
            : base(name, prefix)
        {
            MaxFields = 2;
        }
        public override TemplMatchSection Handler(DocX doc, object model, TemplMatchSection m)
        {
            m.RemovePlaceholder();
            if (m.Fields.Length==2)
            {
                // 2 parts: second part is a bool expression in the model for 'delete section'
                m.Expired = TemplModelEntry.Get(model, m.Fields[1]).AsType<bool>();
            }
            return m;
        }
        public override IEnumerable<TemplMatchSection> FindAll(DocX doc, TemplRegex rxp)
        {
            // Expecting only 1 match per section
            return TemplMatchSection.Find(rxp, doc.GetSections(), 1);
        }
    }
    public class TemplTableModule : TemplModule<TemplMatchTable>
    {
        public TemplTableModule(string name, string prefix)
            : base(name, prefix)
        {
            MaxFields = 2;
        }
        public override TemplMatchTable Handler(DocX doc, object model, TemplMatchTable m)
        {
            m.RemovePlaceholder();
            if (m.Fields.Length==2)
            {
                // 2 parts: second part is a bool expression in the model for 'delete table'
                m.Expired = TemplModelEntry.Get(model, m.Fields[1]).AsType<bool>();
            }
            return m;
        }
        public override IEnumerable<TemplMatchTable> FindAll(DocX doc, TemplRegex rxp)
        {
            // Expecting only 1 match per table
            return TemplMatchTable.Find(rxp, doc.Tables, maxPerTable:1);
        }
    }
    public class TemplRepeatingRowModule : TemplModule<TemplMatchTable>
    {
        public TemplRepeatingRowModule(string name, string prefix)
            : base(name, prefix)
        { }
        public override TemplMatchTable Handler(DocX doc, object model, TemplMatchTable m)
        {
            var e = TemplModelEntry.Get(model, m.Body);
            m.Validate();
            var idx = m.RowIndex;
            m.RemovePlaceholder();
            foreach (var key in e.ToStringKeys())
            {
                var r = m.Table.InsertRow(m.Row, ++idx);
                new TemplSubcollectionModule().BuildFromScope(doc, model, r.Paragraphs, $"{m.Body}[{key}]");
            }
            m.Row.Remove();
            m.Removed = true;
            return m;
        }
        public override IEnumerable<TemplMatchTable> FindAll(DocX doc, TemplRegex rxp)
        {
            // Expecting only 1 match per row
            return TemplMatchTable.Find(rxp, doc.Tables, maxPerRow:1).Reverse();
        }
    }
    public class TemplRepeatingCellModule : TemplModule<TemplMatchTable>
    {
        public TemplRepeatingCellModule(string name, string prefix)
            : base(name, prefix)
        { }
        public override TemplMatchTable Handler(DocX doc, object model, TemplMatchTable m)
        {
            var width = m.Table.Rows.First().Cells.Count;
            var keys = TemplModelEntry.Get(model, m.Body).ToStringKeys();
            var nrows = keys.Count() / width + 1;
            m.Validate();
            m.RemovePlaceholder();
            for (int n=0; n< width; n++)
            {
                if (n != m.CellIndex)
                {
                    CellCopy(m.Cell, m.Row.Cells[n]);
                }
            }
            Row row = m.Row;
            for (int keyIdx = 0; keyIdx < nrows*width; keyIdx++)
            {
                if (keyIdx % width == 0)
                {
                    row = m.Table.InsertRow(m.Row, m.RowIndex + keyIdx/width + 1);
                }
                Cell cell = row.Cells[keyIdx % width];
                if (keyIdx < keys.Count())
                {
                    new TemplSubcollectionModule().BuildFromScope(doc, model, cell.Paragraphs, $"{m.Body}[{keys[keyIdx]}]");
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
        public override IEnumerable<TemplMatchTable> FindAll(DocX doc, TemplRegex rxp)
        {
            // Expecting only 1 match per row (yes really, per row)
            return TemplMatchTable.Find(rxp, doc.Tables, maxPerRow:1).Reverse();
        }
    }
    public class TemplRepeatingTextModule : TemplModule<TemplMatchText>
    {
        public TemplRepeatingTextModule(string name, string prefix)
            : base(name, prefix)
        { }
        public override TemplMatchText Handler(DocX doc, object model, TemplMatchText m)
        {
            var e = TemplModelEntry.Get(model, m.Body);
            m.RemovePlaceholder();
            foreach (var key in e.ToStringKeys())
            {
                var p = m.Paragraph.InsertParagraphAfterSelf(m.Paragraph);
                new TemplSubcollectionModule().BuildFromScope(doc, model, new Paragraph[] { p }, $"{m.Body}[{key}]");
            }
            m.Paragraph.Remove(false);
            m.Removed = true;
            return m;
        }
        public override IEnumerable<TemplMatchText> FindAll(DocX doc, TemplRegex rxp)
        {
            // Expecting only 1 match per paragraph
            return TemplMatchText.Find(rxp, TemplDoc.Paragraphs(doc), 1);
        }
    }
    public class TemplPictureReplaceModule : TemplModule<TemplMatchPicture>
    {
        public TemplPictureReplaceModule(string name, string prefix)
            : base(name, prefix)
        { }
        public override TemplMatchPicture Handler(DocX doc, object model, TemplMatchPicture m)
        {
            var e = TemplModelEntry.Get(model, m.Body);
            var w = m.Picture.Width;
            // Single picture: add text placeholder, expire the placeholder picture
            if (e.Value is TemplGraphic)
            {
                return m.ToText($"{TemplConst.MatchOpen}{TemplConst.Tag.Picture}{TemplConst.FieldSep}{m.Body}{TemplConst.FieldSep}{w}{TemplConst.MatchClose}");
            }
            // Multiple pictures: add repeating list placeholder, expire the placeholder picture
            if (e.Value is TemplGraphic[] || e.Value is ICollection<TemplGraphic>)
            {
                return m.ToText($"{TemplConst.MatchOpen}{TemplConst.Tag.List}{TemplConst.FieldSep}{m.Body}{TemplConst.MatchClose}{TemplConst.MatchOpen}${TemplConst.FieldSep}{TemplConst.Tag.Picture}{TemplConst.FieldSep}{TemplConst.FieldSep}{w}{TemplConst.MatchClose}");
            }
            throw new InvalidCastException($"Templ: Failed to retrieve picture(s) from the model at path \"{e.Path}\"; its actual type is \"{e.Type}\"");
        }
        public override IEnumerable<TemplMatchPicture> FindAll(DocX doc, TemplRegex rxp)
        {
            return TemplMatchPicture.Find(rxp, TemplDoc.Paragraphs(doc));
        }
    }
    public class TemplPicturePlaceholderModule : TemplModule<TemplMatchText>
    {
        public TemplPicturePlaceholderModule(string name, string prefix)
            : base(name, prefix)
        {
            MaxFields = 2;
        }
        public override TemplMatchText Handler(DocX doc, object model, TemplMatchText m)
        {
            int w = -1;
            if (m.Fields.Length==2)
            {
                // 2 parts: second part is an int for the width
                if (!int.TryParse(m.Fields[1], out w))
                {
                    throw new FormatException($"Templ: Picture {m.Body} has a non-integer value for width ({m.Fields[1]})");
                }
            }
            var e = TemplModelEntry.Get(model, m.Fields[0]);
            //Try as array, as collection, as single
            if (e.Value is TemplGraphic)
            {
                return m.ToPicture((e.Value as TemplGraphic).Load(doc), w);
            }
            if (e.Value is TemplGraphic[])
            {
                return m.ToPictures((e.Value as TemplGraphic[])
                        .Select(g => g.Load(doc)).ToArray(), w);
            }
            if (e.Value is ICollection<TemplGraphic>)
            {
                return m.ToPictures((e.Value as ICollection<TemplGraphic>)
                        .Select(g => g.Load(doc)).ToList(), w);
            }
            throw new InvalidCastException($"Templ: Failed to retrieve picture(s) from the model at path \"{e.Path}\"; its actual type is \"{e.Type}\"");
        }
        public override IEnumerable<TemplMatchText> FindAll(DocX doc, TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, TemplDoc.Paragraphs(doc));
        }
    }
    public class TemplTextModule : TemplModule<TemplMatchText>
    {
        public TemplTextModule(string name, string prefix)
            : base(name, prefix)
        { }
        public override TemplMatchText Handler(DocX doc, object model, TemplMatchText m)
        {
            m.ToText(TemplModelEntry.Get(model, m.Body).ToString());
            return m;
        }
        public override IEnumerable<TemplMatchText> FindAll(DocX doc, TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, TemplDoc.Paragraphs(doc));
        }
    }
    public class TemplTOCModule : TemplModule<TemplMatchText>
    {
        public const TableOfContentsSwitches Switches =
            TableOfContentsSwitches.O | TableOfContentsSwitches.H | TableOfContentsSwitches.Z | TableOfContentsSwitches.U;

        public TemplTOCModule(string name, string prefix)
            : base(name, prefix)
        { }

        /// <summary>
        /// Requires user to open in Word and click "Update table of contents". We cannot auto-populate it here.
        /// </summary>
        /// <param name="m"></param>
        /// <param name="doc"></param>
        /// <param name="model"></param>
        /// <returns></returns>
        public override TemplMatchText Handler(DocX doc, object model, TemplMatchText m)
        {
            if (m.Removed)
            {
                return m;
            }
            doc.InsertTableOfContents(m.Paragraph, m.Body, Switches);
            /* Additional options:
                string headerStyle = null
                int maxIncludeLevel = 3
                int? rightTabPos = null) */
            m.RemovePlaceholder();
            return m;
        }
        public override IEnumerable<TemplMatchText> FindAll(DocX doc, TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, TemplDoc.Paragraphs(doc));
        }
    }
    public class TemplCommentsModule : TemplModule<TemplMatchText>
    {
        public TemplCommentsModule(string name, string prefix)
            : base(name, prefix)
        { }
        public override TemplMatchText Handler(DocX doc, object model, TemplMatchText m)
        {
            m.Expired = true;
            return m;
        }
        public override IEnumerable<TemplMatchText> FindAll(DocX doc, TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, TemplDoc.Paragraphs(doc));
        }
    }
}