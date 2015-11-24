using System;
using System.Collections.Generic;
using System.Linq;
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
        public override void Remove()
        {
            base.Remove();
            foreach (Paragraph p in Section.SectionParagraphs.Skip(1))
            {
                p.Pictures.ForEach(pic => pic.Remove());
                p.FollowingTable?.Remove();
                p.Remove(false);
            }
        }
    }
    public class TemplMatchTable : TemplMatchPara
    {
        public Table Table;
        public Row Row => ((RowIndex>=0 && RowIndex<Table?.RowCount) ? Table?.Rows[RowIndex] : null);
        public int RowIndex = -1;
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
        public override void Remove()
        {
            base.Remove();
            Table?.Remove();
        }
        public void Validate()
        {
            if (RowIndex < 0 || RowIndex > Table.RowCount)
            {
                throw new IndexOutOfRangeException($"Templ: Table match \"{Placeholder}\", row index out of bounds ({RowIndex}, Table has {Table.RowCount})");
            }
            if (Table==null)
            {
                throw new NullReferenceException($"Templ: Table match \"{Placeholder}\", Table ref is null");
            }
            if (Row==null)
            {
                throw new NullReferenceException($"Templ: Table match \"{Placeholder}\", Row ref is null");
            }
            if (Cell==null)
            {
                throw new NullReferenceException($"Templ: Table match \"{Placeholder}\", Cell ref is null");
            }
        }
    }
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