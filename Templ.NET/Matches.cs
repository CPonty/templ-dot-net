using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using Novacode;

namespace TemplNET
{
    /// <summary>
    /// Represents a match on a placeholder substring in some arbitrary string
    /// </summary>
    public class TemplMatchString : TemplMatch
    {
        public static IEnumerable<TemplMatchString> Find(TemplRegex rxp, string s)
        {
            return Find<TemplMatchString>(rxp, s);
        }
    }

    /// <summary>
    /// Represents a paragraph, located by matching a placeholder string in the scope of this Paragraph
    /// </summary>
    public class TemplMatchText : TemplMatchPara
    {
        public static IEnumerable<TemplMatchText> Find(TemplRegex rxp, IEnumerable<Paragraph> paragraphs, 
            uint maxPerParagraph = TemplConst.MaxMatchesPerScope) 
        {
            return paragraphs.SelectMany(p => Find(rxp, p).Take((int)maxPerParagraph));
        }
        public static IEnumerable<TemplMatchText> Find(TemplRegex rxp, Paragraph p) 
        {
            return Find<TemplMatchText>(rxp, p);
        }
        public TemplMatchText ToText(string text) => base.ToText<TemplMatchText>(text);
        public TemplMatchText ToPicture(TemplGraphic g, int w) => base.ToPicture<TemplMatchText>(g, w);
        public TemplMatchText ToPictures(TemplGraphic[] g, int w) => base.ToPictures<TemplMatchText>(g, w);
        public TemplMatchText ToPictures(ICollection<TemplGraphic> g, int w) => base.ToPictures<TemplMatchText>(g, w);
    }

    /// <summary>
    /// Represents a section, located by matching a placeholder string in the scope of this section's paragraph
    /// </summary>
    public class TemplMatchSection : TemplMatchPara
    {
        public Section Section;

        public static IEnumerable<TemplMatchSection> Find(TemplRegex rxp, IEnumerable<Section> sections, 
            uint maxPerSection = TemplConst.MaxMatchesPerScope) 
        {
            return sections.SelectMany(sec => Find(rxp, sec).Take((int)maxPerSection));
        }
        public static IEnumerable<TemplMatchSection> Find(TemplRegex rxp, Section sec) 
        {
            return Find<TemplMatchSection>(rxp, sec.SectionParagraphs).Select(m =>
            {
                m.Section = sec;
                return m;
            });
        }
        /// <summary>
        /// Removes the placeholder, plus all *contents* of the section. 
        /// </summary>
        /// Does not Remove() the DocX section object itself - we found this can be problematic (sometimes outputs corrupt document)
        public override void Remove()
        {
            if (Removed)
            {
                return;
            }
            foreach (Paragraph p in Section.SectionParagraphs)
            {
                try
                {
                    p.Pictures?.ForEach(pic => pic.Remove());
                    p.FollowingTable?.Remove();
                    p.Remove(trackChanges: false);
                }
                catch
                {
                }
            }
        }
    }

    /// <summary>
    /// Represents a table cell, located by matching a placeholder string in the scope of this table's paragraphs
    /// </summary>
    public class TemplMatchTable : TemplMatchPara
    {
        public Table Table;
        /// <summary>
        /// Gets the row object at the stored row index
        /// </summary>
        public Row Row => ((RowIndex >= 0 && RowIndex < Table?.RowCount) ? Table?.Rows[RowIndex] : null);
        public int RowIndex = -1;
        /// <summary>
        /// Gets the cell object at the stored row,cell index
        /// </summary>
        public Cell Cell => ((CellIndex>=0 && CellIndex<Row?.Cells?.Count) ? Row?.Cells[CellIndex] : null);
        public int CellIndex = -1;

        public static IEnumerable<TemplMatchTable> Find(TemplRegex rxp, IEnumerable<Table> tables, 
            uint maxPerTable = TemplConst.MaxMatchesPerScope,
            uint maxPerRow = TemplConst.MaxMatchesPerScope,
            uint maxPerCell = TemplConst.MaxMatchesPerScope)
        {
            return tables.SelectMany(t => Find(rxp, t, maxPerTable, maxPerRow, maxPerCell));
        }
        public static IEnumerable<TemplMatchTable> Find(TemplRegex rxp, Table t, 
            uint maxPerTable = TemplConst.MaxMatchesPerScope,
            uint maxPerRow = TemplConst.MaxMatchesPerScope,
            uint maxPerCell = TemplConst.MaxMatchesPerScope)
        {
            return Enumerable.Range(0,t.RowCount).SelectMany(rowIdx => Find(rxp, t, rowIdx, maxPerRow, maxPerCell)).Take((int)maxPerTable);
        }
        public static IEnumerable<TemplMatchTable> Find(TemplRegex rxp, Table t, int rowIdx,
            uint maxPerRow = TemplConst.MaxMatchesPerScope,
            uint maxPerCell = TemplConst.MaxMatchesPerScope)
        {
            return Enumerable.Range(0,t.Rows[rowIdx].Cells.Count).SelectMany(cellIdx => Find(rxp, t, rowIdx, cellIdx, maxPerCell)).Take((int)maxPerRow);
        }
        public static IEnumerable<TemplMatchTable> Find(TemplRegex rxp, Table t, int rowIdx, int cellIdx,
            uint maxPerCell = TemplConst.MaxMatchesPerScope)
        {
            return Find<TemplMatchTable>(rxp, t.Rows[rowIdx].Cells[cellIdx].Paragraphs).Select(m =>
            {
                m.Table = t;
                m.RowIndex = rowIdx;
                m.CellIndex = cellIdx;
                return m;
            }).Take((int)maxPerCell);
        }
        /// <summary>
        /// Removes the matched Table from the document.
        /// </summary>
        public override void Remove()
        {
            base.Remove();
            Table?.Remove();
        }
        /// <summary>
        /// Verifies that the row,cell index refers to a valid cell in the table
        /// </summary>
        public void Validate()
        {
            if (RowIndex < 0 || RowIndex > Table.RowCount)
            {
                throw new IndexOutOfRangeException($"Templ: Table match \"{Placeholder}\", row index out of bounds ({RowIndex}, Table has {Table.RowCount})");
            }
            if (Table == null)
            {
                throw new NullReferenceException($"Templ: Table match \"{Placeholder}\", Table ref is null");
            }
            if (Row == null)
            {
                throw new NullReferenceException($"Templ: Table match \"{Placeholder}\", Row ref is null");
            }
            if (Cell == null)
            {
                throw new NullReferenceException($"Templ: Table match \"{Placeholder}\", Cell ref is null");
            }
        }
    }

    /// <summary>
    /// Represents a picture, located by matching a placeholder string in the scope of this picture's Description field
    /// </summary>
    public class TemplMatchPicture : TemplMatchText
    {
        public Picture Picture;
        public int Width;

        public static IEnumerable<TemplMatchPicture> Find(TemplRegex rxp, IEnumerable<Paragraph> paragraphs) 
        {
            return paragraphs.SelectMany(p => Find(rxp, p));
        }
        public static new IEnumerable<TemplMatchPicture> Find(TemplRegex rxp, Paragraph p) 
        {
            return p.Pictures
                .Where(     pic => pic.Description != null)
                .SelectMany(pic => Find(rxp, p, pic));
        }
        private static IEnumerable<TemplMatchPicture> Find(TemplRegex rxp, Paragraph p, Picture pic) 
        {
            return Find<TemplMatchPicture>(rxp, pic.Description).Select(m => {
                m.Picture = pic;
                m.Width = pic.Width;
                m.Paragraph = p;
                return m;
            });
        }
        /// <summary>
        /// Inserts text at the *beginning* of the Picture's Paragraph. Picture is removed.
        /// </summary>
        public new TemplMatchPicture ToText(string text)
        {
            Paragraph.InsertText(0, text, false);
            Expired = true;
            return this;
        }
        /// <summary>
        /// Sets the string in the Picture's Description property.
        /// </summary>
        public TemplMatchPicture SetDescription(string text)
        {
            if (RemovedPlaceholder || Expired || Removed)
            {
                return this;
            }
            Picture.Description = text;
            return this;
        }
        /// <summary>
        /// Inserts a picture of a specific pixel width. Aspect ratio is maintained. Matched Picture is removed.
        /// </summary>
        private void InsertPicture(TemplGraphic graphic, int width)
        {
            if (RemovedPlaceholder || Expired || Removed)
            {
                return;
            }
            var newPic = graphic.Picture((uint)width);
            Paragraph.InsertPicture(newPic);
            Paragraph.Alignment = (graphic.Alignment ?? Paragraph.Alignment);
            RemovePlaceholder();
            Picture = newPic;
        }
        /// <summary>
        /// Removes the matched Picture from the document.
        /// </summary>
        public override void RemovePlaceholder()
        {
            if (RemovedPlaceholder)
            {
                return;
            }
            try { Picture.Remove(); }
            catch (InvalidOperationException) { } 
            RemovedPlaceholder = true;
            Removed = true;
        }
    }

    /// <summary>
    /// Represents a hyperlink, located by matching a placeholder string in the scope of this hyperlink's Uri
    /// </summary>
    public class TemplMatchHyperlink : TemplMatchText
    {
        public Hyperlink Hyperlink;
        private static string UrlString(Hyperlink hl) => WebUtility.UrlDecode(hl?.Uri?.OriginalString ?? "");

        public static IEnumerable<TemplMatchHyperlink> Find(TemplRegex rxp, IEnumerable<Paragraph> paragraphs) 
        {
            return paragraphs.SelectMany(p => Find(rxp, p));
        }
        public static new IEnumerable<TemplMatchHyperlink> Find(TemplRegex rxp, Paragraph p) 
        {
            return p.Hyperlinks
                .Where(     hl => UrlString(hl).Length > 0)
                .SelectMany(hl => Find(rxp, p, hl));
        }
        private static IEnumerable<TemplMatchHyperlink> Find(TemplRegex rxp, Paragraph p, Hyperlink hl) 
        {
            return Find<TemplMatchHyperlink>(rxp, UrlString(hl)).Select(m => {
                m.Hyperlink= hl;
                m.Paragraph = p;
                return m;
            });
        }

        /// <summary>
        /// Sets the Hyperlink's display text
        /// </summary>
        public TemplMatchHyperlink SetText(string text)
        {
            if (!Removed)
            {
                Hyperlink.Text = text;
            }
            return this;
        }

        /// <summary>
        /// Sets the Hyperlink's Url
        /// </summary>
        public TemplMatchHyperlink SetUrl(string url)
        {
            if (!Removed)
            {
                Hyperlink.Uri = new Uri(url);
                RemovedPlaceholder = true;
            }
            return this;
        }

        /// <summary>
        /// Removes the placeholder from the hyperlink's Url
        /// </summary>
        /// 
        public override void RemovePlaceholder()
        {
            Hyperlink.Uri = new Uri(UrlString(Hyperlink).Replace(Placeholder, ""));
            RemovedPlaceholder = true;
        }

        /// <summary>
        /// Removes the matched Hyperlink from the document.
        /// </summary>
        public override void Remove()
        {
            if (Removed)
            {
                return;
            }
            Hyperlink.Remove();
            RemovedPlaceholder = true;
            Removed = true;
        }
    }
}