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

        public TemplSubcollectionModule(TemplBuilder docBuilder, string name = "Subcollection", string prefix = "$")
            : base(docBuilder, name, prefix)
        {
            MinFields = 2;
            MaxFields = 99;
        }

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
                return (m as TemplMatchPicture).SetDescription(ParentPath(m));
            }
            return m.ToText(ParentPath(m));
        }
        public override IEnumerable<TemplMatchText> FindAll(TemplRegex rxp)
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
        public TemplSectionModule(TemplBuilder docBuilder, string name, string prefix)
            : base(docBuilder, name, prefix)
        {
            MaxFields = 2;
        }
        public override TemplMatchSection Handler(TemplMatchSection m)
        {
            m.RemovePlaceholder();
            if (m.Fields.Length==2)
            {
                // 2 parts: second part is a bool expression in the model for 'delete section'
                m.Expired = TemplModelEntry.Get(Model, m.Fields[1]).AsType<bool>();
            }
            return m;
        }
        public override IEnumerable<TemplMatchSection> FindAll(TemplRegex rxp)
        {
            // Expecting only 1 match per section
            return TemplMatchSection.Find(rxp, Doc.GetSections(), 1);
        }
    }
    public class TemplTableModule : TemplModule<TemplMatchTable>
    {
        public TemplTableModule(TemplBuilder docBuilder, string name, string prefix)
            : base(docBuilder, name, prefix)
        {
            MaxFields = 2;
        }
        public override TemplMatchTable Handler(TemplMatchTable m)
        {
            m.RemovePlaceholder();
            if (m.Fields.Length==2)
            {
                // 2 parts: second part is a bool expression in the model for 'delete table'
                m.Expired = TemplModelEntry.Get(Model, m.Fields[1]).AsType<bool>();
            }
            return m;
        }
        public override IEnumerable<TemplMatchTable> FindAll(TemplRegex rxp)
        {
            // Expecting only 1 match per table
            return TemplMatchTable.Find(rxp, Doc.Tables, maxPerTable:1);
        }
    }
    public class TemplRepeatingRowModule : TemplModule<TemplMatchTable>
    {
        public TemplRepeatingRowModule(TemplBuilder docBuilder, string name, string prefix)
            : base(docBuilder, name, prefix)
        { }
        public override TemplMatchTable Handler(TemplMatchTable m)
        {
            var e = TemplModelEntry.Get(Model, m.Body);
            m.Validate();
            var idx = m.RowIndex;
            m.RemovePlaceholder();
            foreach (var key in e.ToStringKeys())
            {
                var r = m.Table.InsertRow(m.Row, ++idx);
                new TemplSubcollectionModule(DocBuilder).BuildFromScope(r.Paragraphs, $"{m.Body}[{key}]");
            }
            m.Row.Remove();
            m.Removed = true;
            return m;
        }
        public override IEnumerable<TemplMatchTable> FindAll(TemplRegex rxp)
        {
            // Expecting only 1 match per row
            return TemplMatchTable.Find(rxp, Doc.Tables, maxPerRow:1).Reverse();
        }
    }
    public class TemplRepeatingCellModule : TemplModule<TemplMatchTable>
    {
        public TemplRepeatingCellModule(TemplBuilder docBuilder, string name, string prefix)
            : base(docBuilder, name, prefix)
        { }
        public override TemplMatchTable Handler(TemplMatchTable m)
        {
            var width = m.Table.Rows.First().Cells.Count;
            var keys = TemplModelEntry.Get(Model, m.Body).ToStringKeys();
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
                    new TemplSubcollectionModule(DocBuilder).BuildFromScope(cell.Paragraphs, $"{m.Body}[{keys[keyIdx]}]");
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
        public override IEnumerable<TemplMatchTable> FindAll(TemplRegex rxp)
        {
            // Expecting only 1 match per row (yes really, per row)
            return TemplMatchTable.Find(rxp, Doc.Tables, maxPerRow:1).Reverse();
        }
    }
    public class TemplRepeatingTextModule : TemplModule<TemplMatchText>
    {
        public TemplRepeatingTextModule(TemplBuilder docBuilder, string name, string prefix)
            : base(docBuilder, name, prefix)
        { }
        public override TemplMatchText Handler(TemplMatchText m)
        {
            var e = TemplModelEntry.Get(Model, m.Body);
            m.RemovePlaceholder();
            foreach (var key in e.ToStringKeys())
            {
                var p = m.Paragraph.InsertParagraphAfterSelf(m.Paragraph);
                new TemplSubcollectionModule(DocBuilder).BuildFromScope(new Paragraph[] { p }, $"{m.Body}[{key}]");
            }
            m.Paragraph.Remove(false);
            m.Removed = true;
            return m;
        }
        public override IEnumerable<TemplMatchText> FindAll(TemplRegex rxp)
        {
            // Expecting only 1 match per paragraph
            return TemplMatchText.Find(rxp, Doc.Paragraphs, 1);
        }
    }
    public class TemplPictureReplaceModule : TemplModule<TemplMatchPicture>
    {
        public TemplPictureReplaceModule(TemplBuilder docBuilder, string name, string prefix)
            : base(docBuilder, name, prefix)
        { }
        public override TemplMatchPicture Handler(TemplMatchPicture m)
        {
            var e = TemplModelEntry.Get(Model, m.Body);
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
        public override IEnumerable<TemplMatchPicture> FindAll(TemplRegex rxp)
        {
            return TemplMatchPicture.Find(rxp, Doc.Paragraphs);
        }
    }
    public class TemplPicturePlaceholderModule : TemplModule<TemplMatchText>
    {
        public TemplPicturePlaceholderModule(TemplBuilder docBuilder, string name, string prefix)
            : base(docBuilder, name, prefix)
        {
            MaxFields = 2;
        }
        public override TemplMatchText Handler(TemplMatchText m)
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
            var e = TemplModelEntry.Get(Model, m.Fields[0]);
            //Try as array, as collection, as single
            if (e.Value is TemplGraphic)
            {
                return m.ToPicture((e.Value as TemplGraphic).Load(Doc), w);
            }
            if (e.Value is TemplGraphic[])
            {
                return m.ToPictures((e.Value as TemplGraphic[])
                        .Select(g => g.Load(Doc)).ToArray(), w);
            }
            if (e.Value is ICollection<TemplGraphic>)
            {
                return m.ToPictures((e.Value as ICollection<TemplGraphic>)
                        .Select(g => g.Load(Doc)).ToList(), w);
            }
            throw new InvalidCastException($"Templ: Failed to retrieve picture(s) from the model at path \"{e.Path}\"; its actual type is \"{e.Type}\"");
        }
        public override IEnumerable<TemplMatchText> FindAll(TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, Doc.Paragraphs);
        }
    }
    public class TemplTextModule : TemplModule<TemplMatchText>
    {
        public TemplTextModule(TemplBuilder docBuilder, string name, string prefix)
            : base(docBuilder, name, prefix)
        { }
        public override TemplMatchText Handler(TemplMatchText m)
        {
            m.ToText(TemplModelEntry.Get(Model, m.Body).ToString());
            return m;
        }
        public override IEnumerable<TemplMatchText> FindAll(TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, Doc.Paragraphs);
        }
    }
    public class TemplTOCModule : TemplModule<TemplMatchText>
    {
        public const TableOfContentsSwitches Switches =
            TableOfContentsSwitches.O | TableOfContentsSwitches.H | TableOfContentsSwitches.Z | TableOfContentsSwitches.U;

        public TemplTOCModule(TemplBuilder docBuilder, string name, string prefix)
            : base(docBuilder, name, prefix)
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
            Doc.InsertTableOfContents(m.Paragraph, m.Body, Switches);
            /* Additional options:
                string headerStyle = null
                int maxIncludeLevel = 3
                int? rightTabPos = null) */
            m.RemovePlaceholder();
            return m;
        }
        public override IEnumerable<TemplMatchText> FindAll(TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, Doc.Paragraphs);
        }
    }
    public class TemplCommentsModule : TemplModule<TemplMatchText>
    {
        public TemplCommentsModule(TemplBuilder docBuilder, string name, string prefix)
            : base(docBuilder, name, prefix)
        { }
        public override TemplMatchText Handler(TemplMatchText m)
        {
            m.Expired = true;
            return m;
        }
        public override IEnumerable<TemplMatchText> FindAll(TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, Doc.Paragraphs);
        }
    }
}