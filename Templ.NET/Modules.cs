using System;
using System.Collections.Generic;
using System.Linq;
using Novacode;

namespace TemplNET
{
    /// <summary>
    ///This is a utility module, instantiated  by others to handle a sub-scope of the document.
    ///<para/>If you add it to <see cref="Templ.ActiveModules"/>, the default <see cref="TemplModule.Build"/>it won't do anything.
    /// </summary>
    /// This module is to be be instantiated and used within another module's handler.
    /// Its purpose is to allow a scope of the document (e.g. a table row) to be copied multiple times.
    ///
    /// Each copy of the scope is related to an Entry from the Model; most likely an element of a collection/array.
    /// Where any placeholders in the scope contain a model path, the template should use the relative path from the 'parent' Entry in the model.
    /// This module will transform the relative path to absolute.
    ///    
    /// Format: {$:prefix:1..n}
    ///    e.g: {$:txt:Name} // generates {txt:collection[x].Name}
    ///    e.g: {$:txt:}     // generates {txt:collection[x]}
    ///    
    ///prefix=  Any placeholder prefix
    ///  1..n=  Any placeholder fields
    /// <example>
    /// <code>
    /// // Here the Paragraph 'p' is the Nth in a series of copies, 
    /// //  corresponding to some collection called 'objects' in the model
    /// 
    /// int n;
    /// Paragraph p;
    ///
    /// new TemplCollectionModule().BuildFromScope(document, model, new Paragraph[]{ p }, "objects["+n+"]");
    ///
    /// // before: 'p' contained "{$:txt:Name}"
    /// // after:  'p' contains  "{txt:objects["+n+"].Name}"
    /// </code>
    /// <code>
    /// // Here the collection element is a string, 
    /// //  so we want to instert it directly
    /// 
    /// int n;
    /// Paragraph p;
    ///
    /// new TemplCollectionModule().BuildFromScope(document, model, new Paragraph[]{ p }, "strings["+n+"]");
    ///
    /// // before: 'p' contained "{$:txt:}"
    /// // after:  'p' contains  "{txt:strings["+n+"]}"
    /// </code>
    /// </example>
    public class TemplCollectionModule : TemplModule<TemplMatchText>
    {
        private string Path = "";
        private IEnumerable<Paragraph> Paragraphs = new List<Paragraph>();

        public TemplCollectionModule(string name = "Collection", string prefix = TemplConst.Prefix.Collection)
            : base(name, prefix)
        {
            MinFields = 2;
            MaxFields = 99;
        }

        public void BuildFromScope(DocX doc, object model, IEnumerable<Paragraph> paragraphs, string path)
        {
            Path = path;
            Paragraphs = paragraphs;
            Build(doc, model, HandleFailAction.exception);
            Path = "";
            paragraphs = new List<Paragraph>();
        }

        public override TemplMatchText Handler(DocX doc, object model, TemplMatchText m)
        {
            if (m is TemplMatchPicture)
            {
                return (m as TemplMatchPicture).SetDescription(ParentPath(m));
            }
            if (m is TemplMatchHyperlink)
            {
                return (m as TemplMatchHyperlink).SetUrl(ParentPath(m));
            }
            return m.ToText(ParentPath(m));
        }
        public override IEnumerable<TemplMatchText> FindAll(DocX doc, TemplRegex rxp)
        {
            // Get both text and picture matches. Contact is possible because MatchText is MatchPicture's base type.
            return TemplMatchText.Find(rxp, Paragraphs)
                .Concat(TemplMatchPicture.Find(rxp, Paragraphs).Cast<TemplMatchText>())
                .Concat(TemplMatchHyperlink.Find(rxp, Paragraphs).Cast<TemplMatchText>());
        }
        public string ParentPath(TemplMatchText m)
        {
            var fields = m.Fields;
            fields[1] = $"{Path}{(fields[1].Length == 0 || Path.Length == 0?"":".")}{fields[1]}";
            return $"{TemplConst.MatchOpen}{fields.Aggregate((a, b) => $"{a}{TemplConst.FieldSep}{b}") }{TemplConst.MatchClose}";
        }
    }

    /// <summary>
    /// The Section module has 2 purposes:
    /// deleting sections, and providing section matches to <see cref="TemplModule.CustomHandler"/>
    /// </summary>
    /// Format: {sec:path:name}
    ///    e.g: {sec:flags.RemoveSummary:Summary}
    ///    e.g: {sec::CustomSection}
    ///    
    ///  path=  Path in model to a Boolean value. On True, the Section is deleted. Ignored if blank.
    ///  name=  (Optional) Identifies the section in the custom handler.
    public class TemplSectionModule : TemplModule<TemplMatchSection>
    {
        public TemplSectionModule(string name, string prefix = TemplConst.Prefix.Section)
            : base(name, prefix)
        {
            MaxFields = 2;
        }
        public override TemplMatchSection Handler(DocX doc, object model, TemplMatchSection m)
        {
            m.RemovePlaceholder();
            if (m.Fields[0].Length > 0)
            {
                // first field is a bool expression in the model for 'delete section'
                m.Expired = TemplModelEntry.Get(model, m.Fields[0]).AsType<bool>();
            }
            return m;
        }
        public override IEnumerable<TemplMatchSection> FindAll(DocX doc, TemplRegex rxp)
        {
            // Expecting only 1 match per section
            return TemplMatchSection.Find(rxp, doc.GetSections(), 1);
        }
    }

    /// <summary>
    /// The Table module has 2 purposes:
    /// deleting tables, and providing table matches to <see cref="TemplModule.CustomHandler"/>
    /// </summary>
    /// Format: {tab:path:name}
    ///    e.g: {tab:flags.RemoveTables:DataTable}
    ///    e.g: {tab::CustomTable}
    ///    
    ///  path=  Path in model to a Boolean value. On True, the Table is deleted. Ignored if blank.
    ///  name=  (Optional) Identifies the table in the custom handler.
    public class TemplTableModule : TemplModule<TemplMatchTable>
    {
        public TemplTableModule(string name, string prefix = TemplConst.Prefix.Table)
            : base(name, prefix)
        {
            MaxFields = 2;
        }
        public override TemplMatchTable Handler(DocX doc, object model, TemplMatchTable m)
        {
            m.RemovePlaceholder();
            if (m.Fields[0].Length > 0)
            {
                // 2 parts: first part is a bool expression in the model for 'delete table'
                m.Expired = TemplModelEntry.Get(model, m.Fields[0]).AsType<bool>();
            }
            return m;
        }
        public override IEnumerable<TemplMatchTable> FindAll(DocX doc, TemplRegex rxp)
        {
            // Expecting only 1 match per table
            return TemplMatchTable.Find(rxp, doc.Tables, maxPerTable:1);
        }
    }

    /// <summary>
    /// Copies a Matched row N times, where N is the number of items in a collection
    /// </summary>
    /// Format: {row:path}
    ///    e.g: {row:data.people}
    /// 
    ///  path=  Path in model to an Array, Enumerable or Dictionary.
    /// 
    /// If the collection is empty, the matched row is deleted.
    public class TemplRepeatingRowModule : TemplModule<TemplMatchTable>
    {
        public TemplRepeatingRowModule(string name, string prefix = TemplConst.Prefix.Row)
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
                new TemplCollectionModule().BuildFromScope(doc, model, r.Paragraphs, $"{m.Body}[{key}]");
            }
            m.Row.Remove();
            m.Removed = true;
            return m;
        }
        public override IEnumerable<TemplMatchTable> FindAll(DocX doc, TemplRegex rxp)
        {
            // Expecting only 1 match per row
            // Reverse(): If the handler inserts extra rows, we do not want to mess up the indexes of rows above this match.
            //            Processing  bottom-up avoids this issue.
            return TemplMatchTable.Find(rxp, doc.Tables, maxPerRow:1).Reverse();
        }
    }

    /// <summary>
    /// Copies a Matched table cell into a grid pattern N times, where N is the number of items in a collection
    /// </summary>
    /// Format: {cel:path}
    ///    e.g: {cel:data.people}
    /// 
    ///  path=  Path in model to an Array, Enumerable or Dictionary.
    /// 
    /// Cell copying is assumed to be in a grid structure.
    /// The cell is copied from left-to-right, beginning on the matched table row.
    /// Extra table rows are inserted as required.
    /// 
    /// If the collection is empty, the matched row is deleted.
    public class TemplRepeatingCellGridModule : TemplModule<TemplMatchTable>
    {
        public TemplRepeatingCellGridModule(string name, string prefix = TemplConst.Prefix.Cell)
            : base(name, prefix)
        { }
        public override TemplMatchTable Handler(DocX doc, object model, TemplMatchTable m)
        {
            var width = m.Table.Rows.First().Cells.Count;
            var keys = TemplModelEntry.Get(model, m.Body).ToStringKeys();
            var nrows = (int)Math.Ceiling(keys.Count() / (float)width);
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
                    new TemplCollectionModule().BuildFromScope(doc, model, cell.Paragraphs, $"{m.Body}[{keys[keyIdx]}]");
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
        /// <para/>
        /// Optional: Delete all paragraph objects.
        /// Keep in mind, cells with zero paragraphs are considered malformed by Word!
        /// Also keep in mind, you will lose formatting info (e.g. font).
        /// <para/>
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
            // Reverse(): If the handler inserts extra rows, we do not want to mess up the indexes of rows above this match.
            //            Processing  bottom-up avoids this issue.
            return TemplMatchTable.Find(rxp, doc.Tables, maxPerRow:1).Reverse();
        }
    }

    /// <summary>
    /// Copies a Matched paragraph N times, where N is the number of items in a collection
    /// </summary>
    /// Format: {li:path}
    ///    e.g: {li:data.people}
    /// 
    ///  path=  Path in model to an Array, Enumerable or Dictionary.
    /// 
    /// If the collection is empty, the matched paragraph is deleted.
    public class TemplRepeatingTextModule : TemplModule<TemplMatchText>
    {
        public TemplRepeatingTextModule(string name, string prefix = TemplConst.Prefix.List)
            : base(name, prefix)
        { }
        public override TemplMatchText Handler(DocX doc, object model, TemplMatchText m)
        {
            var e = TemplModelEntry.Get(model, m.Body);
            m.RemovePlaceholder();
            foreach (var key in e.ToStringKeys().Reverse())
            {
                var p = m.Paragraph.InsertParagraphAfterSelf(m.Paragraph);
                new TemplCollectionModule().BuildFromScope(doc, model, new Paragraph[] { p }, $"{m.Body}[{key}]");
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

    /// <summary>
    /// Replaces pictures in the document where the Description contains a Picture Placeholder string
    /// </summary>
    /// Format: {pic:path}
    ///    e.g: {pic:images.logoGraphic}
    /// 
    ///  path=  Path in model to a <see cref="TemplGraphic"/>, or Array/Collection of TemplGraphics.
    /// 
    /// ==========================================
    /// The picture is replaced with a placeholder:
    ///         {pic:path:W}
    /// 
    ///     W=  The width of the picture from the document, in pixels
    /// 
    /// The new placeholder will be handled by <see cref="TemplPicturePlaceholderModule"/>.
    public class TemplPictureReplaceModule : TemplModule<TemplMatchPicture>
    {
        public TemplPictureReplaceModule(string name, string prefix = TemplConst.Prefix.Picture)
            : base(name, prefix)
        { }
        public override TemplMatchPicture Handler(DocX doc, object model, TemplMatchPicture m)
        {
            var e = TemplModelEntry.Get(model, m.Body);
            var w = m.Picture.Width;
            // Single picture: add text placeholder, expire the placeholder picture
            if (e.Value is TemplGraphic)
            {
                return m.ToText($"{TemplConst.MatchOpen}{TemplConst.Prefix.Picture}{TemplConst.FieldSep}{m.Body}{TemplConst.FieldSep}{w}{TemplConst.MatchClose}");
            }
            // Multiple pictures: add repeating list placeholder, expire the placeholder picture
            if (e.Value is TemplGraphic[] || e.Value is ICollection<TemplGraphic>)
            {
                return m.ToText($"{TemplConst.MatchOpen}{TemplConst.Prefix.List}{TemplConst.FieldSep}{m.Body}{TemplConst.MatchClose}{TemplConst.MatchOpen}${TemplConst.FieldSep}{TemplConst.Prefix.Picture}{TemplConst.FieldSep}{TemplConst.FieldSep}{w}{TemplConst.MatchClose}");
            }
            throw new InvalidCastException($"Templ: Failed to retrieve picture(s) from the model at path \"{e.Path}\"; its actual type is \"{e.Type}\"");
        }
        public override IEnumerable<TemplMatchPicture> FindAll(DocX doc, TemplRegex rxp)
        {
            return TemplMatchPicture.Find(rxp, TemplDoc.Paragraphs(doc));
        }
    }

    /// <summary>
    /// Updates hyperlinks in the document where the URL contains a URL Placeholder string
    /// </summary>
    /// Format: {url:urlPath}
    /// 
    ///  path=  Path in model to a <see cref="TemplUrl"/>
    /// 
    /// if <see cref="TemplUrl.Url"/> is set, the Hyperlink's Url is updated.
    /// if <see cref="TemplUrl.Text"/> is set, the Hyperlink's Text is updated.
    /// if <see cref="TemplUrl.ToDelete"/> is set, the Hyperlink is removed.
    public class TemplHyperlinkModule : TemplModule<TemplMatchHyperlink>
    {
        public TemplHyperlinkModule(string name, string prefix = TemplConst.Prefix.Hyperlink)
            : base(name, prefix)
        { }
        public override TemplMatchHyperlink Handler(DocX doc, object model, TemplMatchHyperlink m)
        {
            var e = TemplModelEntry.Get(model, m.Body);
            if (e.Value is TemplUrl)
            {
                var url = e.Value as TemplUrl;
                if (url.ToDelete)
                {
                    m.Remove();
                }
                if (url.Text?.Length > 0)
                {
                    m.SetText(url.Text);
                }
                if (url.Url?.Length > 0)
                {
                    m.SetUrl(url.Url);
                }
                return m;
            }
            throw new InvalidCastException($"Templ: Failed to retrieve url from the model at path \"{e.Path}\"; its actual type is \"{e.Type}\"");
        }
        public override IEnumerable<TemplMatchHyperlink> FindAll(DocX doc, TemplRegex rxp)
        {
            return TemplMatchHyperlink.Find(rxp, TemplDoc.Paragraphs(doc));
        }
    }

    /// <summary>
    /// Replaces text placeholders in the Document with Pictures from the model
    /// </summary>
    /// Format: {pic:path:W}
    ///    e.g: {pic:images.logoGraphic}
    ///    e.g: {pic:images.logoGraphic:250}
    /// 
    ///  path=  Path in model to a <see cref="TemplGraphic"/>, or Array/Collection of Graphics.
    ///     W=  (Optional) width in pixels. If none specified, use the width from the Graphic.
    public class TemplPicturePlaceholderModule : TemplModule<TemplMatchText>
    {
        public TemplPicturePlaceholderModule(string name, string prefix = TemplConst.Prefix.Picture)
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

    /// <summary>
    /// Replaces text placeholders in the Document with strings
    /// </summary>
    /// Format: {txt:path}
    ///    e.g: {txt:strings[x]}
    ///    e.g: {txt:Name}
    /// 
    ///  path=  Path in model to a <see cref="string"/>
    public class TemplTextModule : TemplModule<TemplMatchText>
    {
        public TemplTextModule(string name, string prefix = TemplConst.Prefix.Text)
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

    /// <summary>
    /// Removes the paragraph attached to the placeholder
    /// </summary>
    /// Format: {rm:path}
    ///    e.g: {rm:paragraphs[x]}
    ///    e.g: {rm:removeDisclaimer}
    /// 
    ///  path=  Path in model to a <see cref="bool"/>
    public class TemplRemoveModule : TemplModule<TemplMatchText>
    {
        public TemplRemoveModule(string name, string prefix = TemplConst.Prefix.Remove)
            : base(name, prefix)
        { }
        public override TemplMatchText Handler(DocX doc, object model, TemplMatchText m)
        {
            if (TemplModelEntry.Get(model, m.Body).AsType<bool>())
            {
                m.Paragraph.Remove(false);
                m.Removed = true;
                m.Expired = true;
            }
            return m;
        }
        public override IEnumerable<TemplMatchText> FindAll(DocX doc, TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, TemplDoc.Paragraphs(doc), 1);
        }
    }

    /// <summary>
    /// Replaces text placeholders in the Document with a Table of Contents object
    /// </summary>
    /// Format: {toc:title}
    ///    e.g: {toc:Contents}
    /// 
    /// title=  String to use for the title section
    /// 
    /// Note that a user must open the document in Microsoft Word for the object's content to auto-generate.
    public class TemplTOCModule : TemplModule<TemplMatchText>
    {
        public const TableOfContentsSwitches Switches =
            TableOfContentsSwitches.O | TableOfContentsSwitches.H | TableOfContentsSwitches.Z | TableOfContentsSwitches.U;

        public TemplTOCModule(string name, string prefix = TemplConst.Prefix.Contents)
            : base(name, prefix)
        { }

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
            m.Expired = true;
            return m;
        }
        public override IEnumerable<TemplMatchText> FindAll(DocX doc, TemplRegex rxp)
        {
            return TemplMatchText.Find(rxp, TemplDoc.Paragraphs(doc));
        }
    }

    /// <summary>
    /// Removes comments from the Document
    /// </summary>
    /// Format: {!:text}
    ///    e.g: {!:This section of the document is blank; user to enter text}
    /// 
    ///  text=  Any raw string
    public class TemplCommentsModule : TemplModule<TemplMatchText>
    {
        public TemplCommentsModule(string name, string prefix = TemplConst.Prefix.Comment)
            : base(name, prefix)
        {
            MaxFields = 99;
        }
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