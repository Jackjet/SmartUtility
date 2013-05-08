namespace Smart.Utility {
    using System;
	using System.Collections.Generic;
	using System.Drawing;
	using System.Drawing.Imaging;

    public class Thumbnails {

        public Size SingleImage { get; set; }

        public int ColumnCount { get; set; }

        public int Padding { get; set; }

        public Color BackGroundColor { get; set; }

        public Thumbnails()
        {
            SingleImage = new Size(64, 64);
            ColumnCount = 3;
            Padding = 0;
            BackGroundColor = Color.Red;
        }

        public Image GetRectangleImage(List<Image> images)
        {
            Bitmap bitMap = new Bitmap(
                ColumnCount * SingleImage.Width + ColumnCount * Padding,
                images.Count % ColumnCount == 0 ?
                (images.Count / ColumnCount * SingleImage.Height + ColumnCount * Padding) :
                (images.Count / ColumnCount * SingleImage.Height + SingleImage.Height + ColumnCount * Padding));

            Graphics grp = Graphics.FromImage(bitMap);
            Pen pen = new Pen(BackGroundColor);
            pen.Width = bitMap.Height;
            grp.DrawRectangle(pen,
                new Rectangle(new Point(0, 0),
                new Size(bitMap.Width, bitMap.Height)));

            for (int i = 0; i < images.Count; i++)
            {
                Image.GetThumbnailImageAbort call = new Image.GetThumbnailImageAbort(GetThumbnailImageAbort);

                var thumbnailImage = images[i].GetThumbnailImage(SingleImage.Width,
                    SingleImage.Height,
                    call,
                    IntPtr.Zero);

                var point = new Point(
                    (i % ColumnCount * SingleImage.Width),
                    (i / ColumnCount * SingleImage.Height));

                point.Offset(new Point(i % ColumnCount * this.Padding, i / ColumnCount * this.Padding));

                grp.DrawImage(thumbnailImage, point);
            }

            grp.Dispose();

            return bitMap;
        }

        private bool GetThumbnailImageAbort()
        {
            return true;
        }
    }
}
