using System;
namespace Markdown2Openxml
{
    public class ImageSize
    {
        public ImageSize(long width, long height){
            this.Width = width;
            this.Height = height;
        }

        public long Width {get; set; }
        public long Height {get; set; }
    }
}
