using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using TemplNET;

namespace TemplTest
{
    public class ContentDemoModel
    {
        public bool removeAppendix = true;
        public bool removeExtraParagraph = true;
        public string[] list = new string[]
        {
            "Item 1",
            "Item 2",
            "Item 3",
            "Item 4"
        };
        public IEnumerable<string> grid = new List<string>()
        {
            "Cell A1",
            "Cell B1",
            "Cell A2",
            "Cell B2"
        };
        public class MyLocation
        {
            public string Name;
            public int Number;
        }
        public MyLocation Location => Locations[0];
        public MyLocation[] Locations = new MyLocation[] {
            new MyLocation() {Name = "Australia", Number = 1 },
            new MyLocation() {Name = "Brazil", Number = 2 },
            new MyLocation() {Name = "Chile", Number = 3 }
        };

        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd/MM/yy}")]
        public DateTime Date => DateTime.Today.Date;

        public TemplUrl link = new TemplUrl() { Url = "http://www.google.com", Text = "Google's Website" };
    }
}
