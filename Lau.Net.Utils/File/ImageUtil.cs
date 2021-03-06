using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;

namespace Lau.Net.Utils
{
    public static class ImageUtil
    {
        public static Image Scale(this Image @this, double ratio)
        {
            if (@this == null) return null;
            int width = Convert.ToInt32(@this.Width * ratio);
            int height = Convert.ToInt32(@this.Height * ratio);
            return @this.Scale(width, height);
        }

        public static Image Scale(this Image @this, int width, int height)
        {
            if (@this == null) return null;

            var r = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(r))
            {
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;

                g.DrawImage(@this, 0, 0, width, height);
            }

            return r;
        }

        /// <summary>
        /// 将字节数组转换成图像对象
        /// </summary>
        /// <param name="bytes">字节数组</param>
        /// <returns>图像对象</returns>
        public static Image BytesToImage(byte[] bytes)
        {
            if (bytes == null)
            {
                throw new ArgumentNullException("bytes");
            }

            Image image = null;
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(bytes, 0, bytes.Length);
                image = Image.FromStream(stream);
            }
            return image;
        }

        /// <summary>
        /// 将图像对象转换为字节数组
        /// </summary>
        /// <param name="image">Image对象</param>
        /// <returns>字节数组</returns>
        public static byte[] ImageToBytes(Image image)
        {
            if (image == null)
            {
                throw new ArgumentNullException("image");
            }

            byte[] bytes = null;
            using (MemoryStream stream = new MemoryStream())
            {
                image.Save(stream, ImageFormat.Jpeg);
                bytes = stream.ToArray();
            }
            return bytes;
        }
    }
}
