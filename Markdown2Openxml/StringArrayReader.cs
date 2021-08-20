using System;
namespace Markdown2Openxml
{
    public class StringArrayReader
    {
        private int pos = 0;
        private string[] lines;

        public StringArrayReader(string[] array)
        {
            this.lines = array;
        }

        public void increasePos(){
            if(endOfLine())return;
            pos += 1;
        }

        public void decreasePos(){
            if(startOfLine())return;
            pos -= 1;
        }

        public string getCurrentString(){
            return lines[pos];
        }

        public int getCurrentPos(){
            return pos;
        }

        public bool endOfLine(){
            return pos == lines.Length;
        }

        public bool startOfLine(){
            return pos == 0;
        }

        public string nextLineString(){
            if(pos + 1 > lines.Length - 1) return null;
            else return lines[pos + 1];
        }

    }
}
