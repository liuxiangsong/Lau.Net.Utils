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
        #region 缩放
        /// <summary>
        /// 图片按比例缩放
        /// </summary>
        /// <param name="image"></param>
        /// <param name="ratio"></param>
        /// <returns></returns>
        public static Image Scale(Image image, double ratio)
        {
            if (image == null) return null;
            int width = Convert.ToInt32(image.Width * ratio);
            int height = Convert.ToInt32(image.Height * ratio);
            return Scale(image, width, height);
        }

        /// <summary>
        /// 图片按固定宽高缩放
        /// </summary>
        /// <param name="image"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <returns></returns>
        public static Image Scale(Image image, int width, int height)
        {
            if (image == null) return null;

            var r = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(r))
            {
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;

                g.DrawImage(image, 0, 0, width, height);
            }

            return r;
        } 
        #endregion

        /// <summary>
        /// 将字节数组转换成图像对象
        /// </summary>
        /// <param name="bytes">字节数组</param>
        /// <returns>图像对象</returns>
        public static Image ToImage(byte[] bytes)
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
        /// 将图片base64字符串转化为图片
        /// </summary>
        /// <param name="base64">可转换成位图的base64字符串</param>
        /// <returns>Image对象</returns>
        public static Image ToImage(string base64)
        {
            var imageBytes = Convert.FromBase64String(base64);
            using (var ms = new MemoryStream(imageBytes, 0, imageBytes.Length))
            {
                ms.Write(imageBytes, 0, imageBytes.Length);
                return Image.FromStream(ms, true);
            }
        }

        /// <summary>
        /// 将图像对象转换为字节数组
        /// </summary>
        /// <param name="image">Image对象</param>
        /// <returns>字节数组</returns>
        public static byte[] ToBytes(Image image)
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

        /// <summary>
        /// 将图片转换成base64字符串
        /// </summary>
        /// <param name="image">需要转换的图片</param>
        /// <returns>base64字符串</returns>
        public static string ToBase64(Image image)
        {
            var bytes = ToBytes(image);
            return Convert.ToBase64String(bytes);
            //using (var memoryStream = new MemoryStream())
            //{
            //    image.Save(memoryStream, image.RawFormat);
            //    var imageBytes = memoryStream.ToArray();
            //    return Convert.ToBase64String(imageBytes);
            //}
        }
    }
}
