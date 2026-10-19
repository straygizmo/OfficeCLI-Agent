using System.Drawing;
using System.Drawing.Drawing2D;

namespace OfficeCliAgent.Core
{
    /// <summary>Ribbon button images drawn at runtime, so no image resources are needed.</summary>
    public static class RibbonIcons
    {
        /// <summary>A speech bubble with three dots. Drawn at 64px; Office scales it to the button size.</summary>
        public static Bitmap Chat()
        {
            const int size = 64;
            var bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            using (var bubble = new GraphicsPath())
            using (var fill = new SolidBrush(Color.White))
            using (var outline = new Pen(Color.FromArgb(0x42, 0x42, 0x42), 4f) { LineJoin = LineJoin.Round })
            using (var dots = new SolidBrush(Color.FromArgb(0x0F, 0x6C, 0xBD)))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                // Rounded body with a tail at the bottom left.
                const float left = 4, top = 6, right = 60, bottom = 44, r = 12;
                bubble.AddArc(left, top, r * 2, r * 2, 180, 90);
                bubble.AddArc(right - r * 2, top, r * 2, r * 2, 270, 90);
                bubble.AddArc(right - r * 2, bottom - r * 2, r * 2, r * 2, 0, 90);
                bubble.AddLine(right - r, bottom, 26, bottom);
                bubble.AddLine(26, bottom, 12, 58);
                bubble.AddLine(12, 58, 15, bottom);
                bubble.AddArc(left, bottom - r * 2, r * 2, r * 2, 90, 90);
                bubble.CloseFigure();

                g.FillPath(fill, bubble);
                g.DrawPath(outline, bubble);

                const float d = 9, cy = (top + bottom) / 2 - d / 2;
                foreach (var cx in new[] { 18f, 32f, 46f })
                    g.FillEllipse(dots, cx - d / 2, cy, d, d);
            }
            return bmp;
        }
    }
}
