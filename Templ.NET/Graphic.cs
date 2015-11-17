using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Novacode;

namespace Templ
{
    /// <summary>
    /// Used to store/load images and generate (scaled) pictures.
    /// </summary>
    public class TemplGraphic
    {
        protected Novacode.Image Image;
        public MemoryStream Data;
        public double Scalar;
        public Alignment? Alignment;
        private bool Loaded = false;

        public TemplGraphic(Stream data, Alignment? align, double scalar = 1.0)
        {
            Data = new MemoryStream();
            data.CopyTo(Data);
            this.Scalar = scalar;
            this.Alignment = align;
        }
        public TemplGraphic(byte[] data, Alignment? align, double scalar = 1.0)
            : this(new MemoryStream(data), align, scalar)
        { }
        public TemplGraphic(string file, Alignment? align, double scalar = 1.0)
            : this(File.ReadAllBytes(file), align, scalar)
        { }

        public TemplGraphic Load(DocX doc)
        {
            if (!Loaded)
            {
                Image = doc.AddImage(Data);
                Loaded = true;
            }
            return this;
        }

        public Picture Picture(double scalar)
        {
            var pic = Image.CreatePicture();
            pic.Width = (int)(pic.Width * (scalar));
            pic.Height = (int)(pic.Height * (scalar));
            return pic;
        }
        public Picture Picture()
        {
            return Picture(Scalar);
        }
        public Picture Picture(uint width)
        {
            var pic = Image.CreatePicture();
            var scaling = ((float)width / pic.Width);
            pic.Height = (int)(pic.Height * scaling);
            pic.Width = (int)(pic.Width * scaling);
            return pic;
        }
    }
}