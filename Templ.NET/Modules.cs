using System;
using System.Collections.Generic;
using System.Linq;
using Novacode;

namespace Templ
{
    public abstract class TemplModule<T> : TemplModule where T : TemplMatchPara, new()
    {
        public static uint MinFields;
        public static uint MaxFields;
        public new Func<T, DocX, object, T> CustomHandler = (m, doc, model) => m;

        public TemplModule(String name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }

        private T CheckFieldCount(T m)
        {
            var l = m.Fields.Length;
            if (l < MinFields)
            {
                throw new FormatException($"Templ: Module \"{GetType()}\" found a placeholder \"{m.Placeholder}\" with too few :-separated fields ({l} vs {MinFields})");
            }
            if (l > MaxFields)
            {
                throw new FormatException($"Templ: Module \"{GetType()}\" found a placeholder \"{m.Placeholder}\" with too many :-separated fields ({l} vs {MaxFields})");
            }
            return m;
        }
        public void BuildFromScope(IEnumerable<T> scope)
        {
            // Note how we are constantly "committing" the changes using ToList().
            // This ensures order is preserved (e.g. all "finds" happen before all "handler"s)
            scope.Select( m => CheckFieldCount(m) ).ToList()
                 .Select( m => Handler(m)).ToList()
                 .Where(  m =>!m.RemoveExpired()).ToList()
                 .Select( m => CustomHandler(m, Doc, Model)).ToList()
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
        private TemplBuilder DocBuilder;
        public DocX Doc => DocBuilder.Doc;
        public object Model => DocBuilder.Model;
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
        public static new uint MinFields = 2;
        public static new uint MaxFields = 99;
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
            m.Fields[1] = $"{Path}{(m.Fields[1].Length == 0 || Path.Length == 0?"":".")}{m.Fields[1]}";
            return $"{{{m.Fields.Aggregate((a, b) => $"{a}:{b}") }}}";
        }
    }

    public class TemplSectionModule : TemplModule<TemplMatchSection>
    {
        public static new uint MaxFields = 2;
        public TemplSectionModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchSection Handler(TemplMatchSection m)
        {
            m.RemovePlaceholder();
            switch (m.Fields.Length)
            {
                case 1:
                    // 1 part: placeholder is just a label, nothing to do.
                    // Users can implement their own section handler to switch off this label.
                    return m;
                case 2:
                    // 2 parts: second part is a bool expression in the model for 'delete section'
                    m.Expired = TemplModelEntry.Get(Model, m.Fields[1]).AsType<bool>();
                    return m;
                default:
                    throw new FormatException($"Templ: Section {m.Body} has a malformed tag body (too many occurrences of :)");
            }
        }
        public override IEnumerable<TemplMatchSection> FindAll(TemplRegex rxp)
        {
            // Expecting only 1 match per section
            return TemplMatchSection.Find(rxp, Doc.GetSections(), 1);
        }
    }
    public class TemplTableModule : TemplModule<TemplMatchTable>
    {
        public static new uint MaxFields = 2;
        public TemplTableModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchTable Handler(TemplMatchTable m)
        {
            m.RemovePlaceholder();
            switch (m.Fields.Length)
            {
                case 1:
                    // 1 part: placeholder name is just a label, nothing to do.
                    // Users can implement their own table handler to switch off this label.
                    return m;
                case 2:
                    // 2 parts: second part is a bool expression in the model for 'delete table'
                    m.Expired = TemplModelEntry.Get(Model, m.Fields[1]).AsType<bool>();
                    return m;
                default:
                    // 3+ parts: placeholder name is malformed
                    throw new FormatException($"Templ: Table \"{m.Body}\" has a malformed tag body (too many occurrences of :)");
            }
        }
        public override IEnumerable<TemplMatchTable> FindAll(TemplRegex rxp)
        {
            // Expecting only 1 match per table
            return TemplMatchTable.Find(rxp, Doc.Tables, 1);
        }
    }
    public class TemplRepeatingRowModule : TemplModule<TemplMatchRow>
    {
        public TemplRepeatingRowModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchRow Handler(TemplMatchRow m)
        {
            var e = TemplModelEntry.Get(Model, m.Body);
            var idx = m.RowIndex;
            if (idx < 0)
            {
                throw new IndexOutOfRangeException($"Templ: Row match does not have an index in a table for match \"{m.Body}\"");
            }
            m.RemovePlaceholder();
            foreach (var key in e.ToStringKeys())
            {
                var r = m.Table.InsertRow(m.Row, ++idx);
                new TemplSubcollectionModule().BuildFromScope(r.Paragraphs, $"{m.Body}[{key}]");
            }
            m.Row.Remove();
            m.Removed = true;
            return m;
        }
        public override IEnumerable<TemplMatchRow> FindAll(TemplRegex rxp)
        {
            // Expecting only 1 match per row
            return TemplMatchRow.Find(rxp, Doc.Tables, 1);
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
            var keys = TemplModelEntry.Get(Model, m.Body).ToStringKeys();
            var nrows = keys.Count() / width + 1;
            if (m.RowIndex < 0)
            {
                throw new IndexOutOfRangeException($"Templ: Cell match does not have a row index in a table for match \"{m.Body}\"");
            }
            if (m.CellIndex < 0)
            {
                throw new IndexOutOfRangeException($"Templ: Cell match does not have a cell index in a table for match \"{m.Body}\"");
            }
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
                    new TemplSubcollectionModule().BuildFromScope(cell.Paragraphs, $"{m.Body}[{keys[keyIdx]}]");
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
            // Expecting only 1 match per cell
            return TemplMatchCell.Find(rxp, Doc.Tables, 1);
        }
    }
    public class TemplRepeatingTextModule : TemplModule<TemplMatchText>
    {
        public TemplRepeatingTextModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchText Handler(TemplMatchText m)
        {
            var e = TemplModelEntry.Get(Model, m.Body);
            m.RemovePlaceholder();
            foreach (var key in e.ToStringKeys())
            {
                var p = m.Paragraph.InsertParagraphAfterSelf(m.Paragraph);
                new TemplSubcollectionModule().BuildFromScope(new Paragraph[] { p }, $"{m.Body}[{key}]");
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
        public TemplPictureReplaceModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchPicture Handler(TemplMatchPicture m)
        {
            if (m.Fields.Length > 1)
            {
                throw new FormatException($"Templ: Template Picture {m.Body} has a malformed tag body (too many occurrences of :)");
            }
            var e = TemplModelEntry.Get(Model, m.Body);
            var w = m.Picture.Width;
            // Single picture: add text placeholder, expire the placeholder picture
            if (e.Value is TemplGraphic)
            {
                return m.ToText($"{{pic:{m.Body}:{w}}}");
            }
            // Multiple pictures: add repeating list placeholder, expire the placeholder picture
            if (e.Value is TemplGraphic[] || e.Value is ICollection<TemplGraphic>)
            {
                return m.ToText($"{{li:{m.Body}}}{{$:pic::{w}}}");
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
        public static new uint MaxFields = 2;
        public TemplPicturePlaceholderModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }
        public override TemplMatchText Handler(TemplMatchText m)
        {
            int w = -1;
            switch (m.Fields.Length)
            {
                case 1:
                    break;
                case 2:
                    // 2 parts: second part is an int for the width
                    if (!int.TryParse(m.Fields[1], out w))
                    {
                        throw new FormatException($"Templ: Picture {m.Body} has a non-integer value for width ({m.Fields[1]})");
                    }
                    break;
                default:
                    // 3+ parts: placeholder body is malformed
                    throw new FormatException($"Templ: Picture {m.Body} has a malformed tag body (too many occurrences of :)");
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
        public TemplTextModule(string name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
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
            return TemplMatchText.Find(rxp, Doc.Paragraphs);
        }
    }
}