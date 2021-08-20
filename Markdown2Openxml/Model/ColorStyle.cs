using System;
namespace Markdown2Openxml.Model
{
    public class ColorStyle
    {
        public ColorStyle(string description, string color)
        {
            this.Description = description;
            this.Color = color;
        }

        public string Description { get; set; }
        public string Color { get; set; }
    }
}
