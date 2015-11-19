using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Novacode;

namespace TemplNET
{
    public class TemplMatchString : TemplMatch
    {
        public static IEnumerable<TemplMatchString> Find(TemplRegex rxp, string s)
        {
            return Find<TemplMatchString>(rxp, s);
        }
    }
    public class TemplMatchText : TemplMatchPara
    {
        public static IEnumerable<TemplMatchText> Find(TemplRegex rxp, IEnumerable<Paragraph> paragraphs) 
        {
            return paragraphs.SelectMany(p => Find(rxp, p));
        }
        public static IEnumerable<TemplMatchText> Find(TemplRegex rxp, IEnumerable<Paragraph> paragraphs, int n) 
        {
            return paragraphs.SelectMany(p => Find(rxp, p).Take(n));
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
    public class TemplMatchSection : TemplMatchPara
    {
        public Section Section;

        public static IEnumerable<TemplMatchSection> Find(TemplRegex rxp, IEnumerable<Section> sections) 
        {
            return sections.SelectMany(sec => Find(rxp, sec));
        }
        public static IEnumerable<TemplMatchSection> Find(TemplRegex rxp, IEnumerable<Section> sections, int n) 
        {
            return sections.SelectMany(sec => Find(rxp, sec).Take(n));
        }
        public static IEnumerable<TemplMatchSection> Find(TemplRegex rxp, Section sec) 
        {
            return Find<TemplMatchSection>(rxp, sec.SectionParagraphs).Select(m =>
            {
                m.Section = sec;
                return m;
            });
        }
        public override void Remove()
        {
            RemovePlaceholder();
            foreach (Paragraph p in Section.SectionParagraphs.Skip(1))
            {
                p.Pictures.ForEach(pic => pic.Remove());
                p.FollowingTable?.Remove();
                p.Remove(false);
            }
            Removed = true;
        }
    }
    public class TemplMatchTable : TemplMatchPara
    {
        public Table Table;

        public static IEnumerable<TemplMatchTable> Find(TemplRegex rxp, IEnumerable<Table> tables, int n) 
        {
            return tables.SelectMany(t => Find(rxp, t).Take(n));
        }
        public static IEnumerable<TemplMatchTable> Find(TemplRegex rxp, IEnumerable<Table> tables) 
        {
            return tables.SelectMany(t => Find(rxp, t));
        }
        public static IEnumerable<TemplMatchTable> Find(TemplRegex rxp, Table t) 
        {
            return Find<TemplMatchTable>(rxp, t.Rows[0].Cells[0].Paragraphs).Select(m =>
            {
                m.Table = t;
                return m;
            });
        }
    }
    public class TemplMatchRow : TemplMatchTable
    {
        public Row Row;
        public int RowIndex;

        public static new IEnumerable<TemplMatchRow> Find(TemplRegex rxp, IEnumerable<Table> tables, int n)
        {
            return tables.SelectMany(t => Find(rxp, t, n));
        }
        public static new IEnumerable<TemplMatchRow> Find(TemplRegex rxp, IEnumerable<Table> tables)
        {
            return tables.SelectMany(t => Find(rxp, t));
        }
        public static new IEnumerable<TemplMatchRow> Find(TemplRegex rxp, Table t)
        {
            return Enumerable.Range(0,t.RowCount).SelectMany(rIdx => FindInRow(rxp, t, rIdx));
        }
        public static IEnumerable<TemplMatchRow> Find(TemplRegex rxp, Table t, int n)
        {
            return Enumerable.Range(0,t.RowCount).SelectMany(rIdx => FindInRow(rxp, t, rIdx).Take(n));
        }

        private static IEnumerable<TemplMatchRow> FindInRow(TemplRegex rxp, Table t, int rowIndex)
        {
            var r = t.Rows[rowIndex];
            return Find<TemplMatchRow>(rxp, r.Cells.SelectMany(c => c.Paragraphs)).Select(m =>
            {
                m.Table = t;
                m.Row = r;
                m.RowIndex = rowIndex;
                return m;
            });
        }
    }
    public class TemplMatchCell : TemplMatchRow
    {
        public Cell Cell;
        public int CellIndex;

        public static new IEnumerable<TemplMatchCell> Find(TemplRegex rxp, IEnumerable<Table> tables)
        {
            return tables.SelectMany(t => Find(rxp, t));
        }
        public static new IEnumerable<TemplMatchCell> Find(TemplRegex rxp, IEnumerable<Table> tables, int n)
        {
            return tables.SelectMany(t => Find(rxp, t, n));
        }
        public static new IEnumerable<TemplMatchCell> Find(TemplRegex rxp, Table t)
        {
            return Enumerable.Range(0,t.RowCount).SelectMany(rowIdx => FindInRow(rxp, t, rowIdx));
        }
        public static new IEnumerable<TemplMatchCell> Find(TemplRegex rxp, Table t, int n)
        {
            return Enumerable.Range(0,t.RowCount).SelectMany(rowIdx => FindInRow(rxp, t, rowIdx, n));
        }

        private static IEnumerable<TemplMatchCell> FindInRow(TemplRegex rxp, Table t, int rowIdx)
        {
            return Enumerable.Range(0,t.Rows[rowIdx].Cells.Count).SelectMany(cellIdx => FindInCell(rxp, t, rowIdx, cellIdx));
        }
        private static IEnumerable<TemplMatchCell> FindInRow(TemplRegex rxp, Table t, int rowIdx, int n)
        {
            return Enumerable.Range(0,t.Rows[rowIdx].Cells.Count).SelectMany(cellIdx => FindInCell(rxp, t, rowIdx, cellIdx).Take(1));
        }

        private static IEnumerable<TemplMatchCell> FindInCell(TemplRegex rxp, Table t, int rowIdx, int cellIdx)
        {
            var r = t.Rows[rowIdx];
            var c = r.Cells[cellIdx];
            return Find<TemplMatchCell>(rxp, c.Paragraphs).Select(m =>
            {
                m.Table = t;
                m.Row = r;
                m.Cell = c;
                return m;
            });
        }
    }
    public class TemplMatchPicture : TemplMatchText
    {
        public Picture Picture;
        public int Width;

        public static new IEnumerable<TemplMatchPicture> Find(TemplRegex rxp, IEnumerable<Paragraph> paragraphs) 
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
        public new TemplMatchPicture ToText(string text)
        {
            Paragraph.InsertText(0, text, false);
            Expired = true;
            return this;
        }
        public TemplMatchPicture SetDescription(string text)
        {
            if (RemovedPlaceholder || Expired || Removed)
            {
                return this;
            }
            Picture.Description = text;
            return this;
        }
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
}