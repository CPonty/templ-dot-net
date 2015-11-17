using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using Novacode;

namespace DocXMVC
{
    public abstract class DocxMatch
    {
        public bool Success = false;
        public bool Expired = false;
        public bool Removed = false;
        protected bool RemovedPlaceholder = false;
        public string Name;
        public DocxRegex Regex;
        public string Placeholder => Regex.Text(Name);
        public Regex Pattern => Regex.Pattern;

        // if someone tries to implement single-result search: .DefaultIfEmpty(new T()).First();
        public static IEnumerable<T> Find<T>(DocxRegex rgx, IEnumerable<Paragraph> paragraphs) where T : DocxMatchPara, new()
        {
            return paragraphs.SelectMany(p => Find<T>(rgx, p));
        }
        public static IEnumerable<T> Find<T>(DocxRegex rgx, Paragraph p) where T : DocxMatchPara, new()
        {
            var P = p;
            var t = P.Text;
            return Find<T>(rgx, p.Text).Select(m =>
            {
                m.Paragraph = p;
                return m;
            });
        }
        public static IEnumerable<T> Find<T>(DocxRegex rgx, string s) where T : DocxMatch, new()
        {
            return rgx.Pattern.Matches(s).Cast<Match>().Where(m => m.Success).Select(m => new T()
            {
                Success = true,
                Regex = rgx,
                Name = m.Groups[1].Value,
            });
        }
    }
    public abstract class DocxMatchPara : DocxMatch
    {
        public Paragraph Paragraph;

        private void InsertText(int idx, string text)
        {
            // Simple. Text at index
            Paragraph.InsertText(idx, text);
        }

        public T ToText<T>(string text) where T : DocxMatchPara, new()
        {
            // Insert Text at each match, then remove placeholder
            Paragraph.ReplaceText(Placeholder, text);
            RemovedPlaceholder = true;
            return (T)this;
        }
        public Paragraph AddBefore<T>(string text) where T : DocxMatchPara, new() => Paragraph.InsertParagraphBeforeSelf(text);
        public Paragraph AddAfter<T>(string text) where T : DocxMatchPara, new() => Paragraph.InsertParagraphAfterSelf(text);

        private void InsertPicture(DocxGraphic graphic, int width)
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
        public T ToPicture<T>(DocxGraphic graphic, int width) where T : DocxMatchPara, new()
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

        public T ToPictures<T>(ICollection<DocxGraphic> graphics, int width) where T : DocxMatchPara, new()
        {
            return ToPictures<T>(graphics.ToArray(), width);
        }

        public T ToPictures<T>(DocxGraphic[] graphics, int width) where T : DocxMatchPara, new()
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

    public class DocxMatchString : DocxMatch
    {
        public static IEnumerable<DocxMatchString> Find(DocxRegex rxp, string s)
        {
            return Find<DocxMatchString>(rxp, s);
        }
    }
    public class DocxMatchText : DocxMatchPara
    {
        public static IEnumerable<DocxMatchText> Find(DocxRegex rxp, IEnumerable<Paragraph> paragraphs) 
        {
            return paragraphs.SelectMany(p => Find(rxp, p));
        }
        public static IEnumerable<DocxMatchText> Find(DocxRegex rxp, Paragraph p) 
        {
            return Find<DocxMatchText>(rxp, p);
        }
        public DocxMatchText ToText(string text) => base.ToText<DocxMatchText>(text);
        public DocxMatchText ToPicture(DocxGraphic g, int w) => base.ToPicture<DocxMatchText>(g, w);
        public DocxMatchText ToPictures(DocxGraphic[] g, int w) => base.ToPictures<DocxMatchText>(g, w);
        public DocxMatchText ToPictures(ICollection<DocxGraphic> g, int w) => base.ToPictures<DocxMatchText>(g, w);
    }
    public class DocxMatchSection : DocxMatchPara
    {
        public Section Section;

        public static IEnumerable<DocxMatchSection> Find(DocxRegex rxp, IEnumerable<Section> sections) 
        {
            return sections.SelectMany(sec => Find(rxp, sec));
        }
        public static IEnumerable<DocxMatchSection> Find(DocxRegex rxp, Section sec) 
        {
            return Find<DocxMatchSection>(rxp, sec.SectionParagraphs).Select(m =>
            {
                m.Section = sec;
                return m;
            });
        }
        public override void Remove()
        {
            RemovePlaceholder();
            //foreach (Paragraph p in Section.SectionParagraphs.Skip(1).Where(p => p.Text.Length > 0))
            foreach (Paragraph p in Section.SectionParagraphs.Skip(1))
            {
                //p.RemoveText(0, p.Text.Length);
                p.Pictures.ForEach(pic => pic.Remove());
                p.FollowingTable?.Remove();
                p.Remove(false);
            }
            Removed = true;
        }
    }
    public class DocxMatchTable : DocxMatchPara
    {
        public Table Table;

        public static IEnumerable<DocxMatchTable> Find(DocxRegex rxp, IEnumerable<Table> tables) 
        {
            return tables.SelectMany(t => Find(rxp, t));
        }
        public static IEnumerable<DocxMatchTable> Find(DocxRegex rxp, Table t) 
        {
            return Find<DocxMatchTable>(rxp, t.Rows[0].Cells[0].Paragraphs).Select(m =>
            {
                m.Table = t;
                return m;
            });
        }
    }
    public class DocxMatchRow : DocxMatchTable
    {
        public Row Row;
        public int RowIndex;

        public static new IEnumerable<DocxMatchRow> Find(DocxRegex rxp, IEnumerable<Table> tables)
        {
            return tables.SelectMany(t => Find(rxp, t));
        }
        public static new IEnumerable<DocxMatchRow> Find(DocxRegex rxp, Table t)
        {
            return Enumerable.Range(0,t.RowCount).SelectMany(rIdx => Find(rxp, t, rIdx));
        }

        public static IEnumerable<DocxMatchRow> Find(DocxRegex rxp, Table t, int rowIndex)
        {
            var r = t.Rows[rowIndex];
            return Find<DocxMatchRow>(rxp, r.Cells.SelectMany(c => c.Paragraphs)).Select(m =>
            {
                m.Table = t;
                m.Row = r;
                m.RowIndex = rowIndex;
                return m;
            });
        }
    }
    public class DocxMatchCell : DocxMatchRow
    {
        public Cell Cell;
        public int CellIndex;

        public static new IEnumerable<DocxMatchCell> Find(DocxRegex rxp, IEnumerable<Table> tables)
        {
            return tables.SelectMany(t => Find(rxp, t));
        }
        public static new IEnumerable<DocxMatchCell> Find(DocxRegex rxp, Table t)
        {
            return Enumerable.Range(0,t.RowCount).SelectMany(rowIdx => Find(rxp, t, rowIdx));
        }

        public static new IEnumerable<DocxMatchCell> Find(DocxRegex rxp, Table t, int rowIdx)
        {
            return Enumerable.Range(0,t.Rows[rowIdx].Cells.Count).SelectMany(cellIdx => Find(rxp, t, rowIdx, cellIdx));
        }

        public static IEnumerable<DocxMatchCell> Find(DocxRegex rxp, Table t, int rowIdx, int cellIdx)
        {
            var r = t.Rows[rowIdx];
            var c = r.Cells[cellIdx];
            return Find<DocxMatchCell>(rxp, c.Paragraphs).Select(m =>
            {
                m.Table = t;
                m.Row = r;
                m.Cell = c;
                return m;
            });
        }
    }
    public class DocxMatchPicture : DocxMatchText
    {
        public Picture Picture;
        public int Width;

        public static new IEnumerable<DocxMatchPicture> Find(DocxRegex rxp, IEnumerable<Paragraph> paragraphs) 
        {
            return paragraphs.SelectMany(p => Find(rxp, p));
        }
        public static new IEnumerable<DocxMatchPicture> Find(DocxRegex rxp, Paragraph p) 
        {
            return p.Pictures
                .Where(     pic => pic.Description != null)
                .SelectMany(pic => Find(rxp, p, pic));
        }
        private static IEnumerable<DocxMatchPicture> Find(DocxRegex rxp, Paragraph p, Picture pic) 
        {
            return Find<DocxMatchPicture>(rxp, pic.Description).Select(m => {
                m.Picture = pic;
                m.Width = pic.Width;
                m.Paragraph = p;
                return m;
            });
        }
        public new DocxMatchPicture ToText(string text)
        {
            Paragraph.InsertText(0, text, false);
            Expired = true;
            return this;
        }
        public DocxMatchPicture SetDescription(string text)
        {
            if (RemovedPlaceholder || Expired || Removed)
            {
                return this;
            }
            Picture.Description = text;
            return this;
        }
        private void InsertPicture(DocxGraphic graphic, int width)
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
        // The below code is perfectly fine, but we have superseded it with text conversion,
        //  and converting back to the new image later.
        // Remove soon, if we don't end up needing it again.
        /*
        public DocxMatchPicture ToPicture(DocxGraphic g)
        {
            if (Removed || RemovedPlaceholder || Expired)
            {
                return this;
            }
            InsertPicture(g, Picture.Width);
            return this;
        }
        public DocxMatchPicture ToPictures(DocxGraphic[] g)
        {
            var width = Picture.Width;
            foreach(var graphic in g.Reverse())
            {
                InsertPicture(graphic, width);
            }
            return this;
        }
        public DocxMatchPicture ToPictures(ICollection<DocxGraphic> g)
        {
            return ToPictures(g.ToArray());
        }
        */
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