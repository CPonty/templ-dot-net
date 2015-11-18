using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Novacode;

namespace TemplNET
{
    public abstract class TemplMatch
    {
        public bool Success = false;
        public bool Expired = false;
        public bool Removed = false;
        protected bool RemovedPlaceholder = false;
        public string Body;
        public string[] Fields => Body.Split(':');
        public TemplRegex Regex;
        public string Placeholder => Regex.Text(Body);
        public Regex Pattern => Regex.Pattern;

        // if someone tries to implement single-result search: .DefaultIfEmpty(new T()).First();
        public static IEnumerable<T> Find<T>(TemplRegex rgx, IEnumerable<Paragraph> paragraphs) where T : TemplMatchPara, new()
        {
            return paragraphs.SelectMany(p => Find<T>(rgx, p));
        }
        public static IEnumerable<T> Find<T>(TemplRegex rgx, IEnumerable<Paragraph> paragraphs, int n) where T : TemplMatchPara, new()
        {
            return paragraphs.SelectMany(p => Find<T>(rgx, p).Take(n));
        }
        public static IEnumerable<T> Find<T>(TemplRegex rgx, Paragraph p) where T : TemplMatchPara, new()
        {
            var P = p;
            var t = P.Text;
            return Find<T>(rgx, p.Text).Select(m =>
            {
                m.Paragraph = p;
                return m;
            });
        }
        public static IEnumerable<T> Find<T>(TemplRegex rgx, string s) where T : TemplMatch, new()
        {
            return rgx.Pattern.Matches(s).Cast<Match>().Where(m => m.Success).Select(m => new T()
            {
                Success = true,
                Regex = rgx,
                Body = m.Groups[1].Value,
            });
        }
    }
    public abstract class TemplMatchPara : TemplMatch
    {
        public Paragraph Paragraph;

        private void InsertText(int idx, string text)
        {
            // Simple. Text at index
            Paragraph.InsertText(idx, text);
        }

        public T ToText<T>(string text) where T : TemplMatchPara, new()
        {
            // Insert Text at each match, then remove placeholder
            Paragraph.ReplaceText(Placeholder, text);
            RemovedPlaceholder = true;
            return (T)this;
        }
        public Paragraph AddBefore<T>(string text) where T : TemplMatchPara, new() => Paragraph.InsertParagraphBeforeSelf(text);
        public Paragraph AddAfter<T>(string text) where T : TemplMatchPara, new() => Paragraph.InsertParagraphAfterSelf(text);

        private void InsertPicture(TemplGraphic graphic, int width)
        {
            // Simple. Picture at index(es). Ensure you load picture first.
            Paragraph.FindAll(Placeholder).ForEach(idx =>
            {
                if (width>0)
                {
                    Paragraph.InsertPicture(graphic.Picture((uint)width), idx + Placeholder.Length);
                }
                else
                {
                    Paragraph.InsertPicture(graphic.Picture(), idx + Placeholder.Length);
                }
                Paragraph.Alignment = (graphic.Alignment ?? Paragraph.Alignment);
            });
        }
        public T ToPicture<T>(TemplGraphic graphic, int width) where T : TemplMatchPara, new()
        {
            // Use InsertPicture at each match, then remove placeholder
            if (Removed || RemovedPlaceholder || Expired)
            {
                return (T)this;
            }
            InsertPicture(graphic,width);
            RemovePlaceholder();
            return (T)this;
        }

        public T ToPictures<T>(ICollection<TemplGraphic> graphics, int width) where T : TemplMatchPara, new()
        {
            return ToPictures<T>(graphics.ToArray(), width);
        }

        public T ToPictures<T>(TemplGraphic[] graphics, int width) where T : TemplMatchPara, new()
        {
            if (Removed || RemovedPlaceholder || Expired)
            {
                return (T)this;
            }
            // Use InsertPicture for each graphic, then remove placeholder.
            // Reverse order as pictures get inserted immediately after the placeholder, 
            //  so each subsequent picture inserted will appear before those already inserted.
            foreach (var graphic in graphics.Reverse())
            {
                InsertPicture(graphic, width);
            }
            RemovePlaceholder();
            return (T)this;
        }

        public virtual void Remove()
        {
            // By default, Remove() simply removes the placeholder. 
            // For more advanced content it will do more.
            if (Removed)
            {
                return;
            }
            RemovePlaceholder();
            Removed = true;
        }

        public virtual void RemovePlaceholder()
        {
            if (RemovedPlaceholder)
            {
                return;
            }
            Paragraph.ReplaceText(Placeholder, string.Empty);
            RemovedPlaceholder = true;
        }

        public bool RemoveExpired()
        {
            if (Expired)
            {
                Remove();
            }
            return Removed;
        }
    }

    /* ---------------------------------------------------------------------- */
    /* ---------------------------------------------------------------------- */
    /* ---------------------------------------------------------------------- */

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