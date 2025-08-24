// WindowCapture.cs
using System;
using System.Drawing;
using System.Drawing.Imaging;

namespace TouchKeyBoard
{
    public static class WindowCapture
    {
        /// <summary>
        /// 根據視窗控制代碼 (HWND) 擷取該視窗的畫面。
        /// </summary>
        /// <param name="hWnd">目標視窗的控制代碼。</param>
        /// <returns>一個包含視窗畫面的 Bitmap 物件，如果失敗則返回 null。</returns>
        public static Bitmap CaptureWindow(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero || !NativeMethods.IsWindowVisible(hWnd))
            {
                return null;
            }

            try
            {
                NativeMethods.GetWindowRect(hWnd, out var rect);
                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;

                if (width <= 0 || height <= 0)
                {
                    return null;
                }

                Bitmap bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
                using (Graphics graphics = Graphics.FromImage(bmp))
                {
                    IntPtr hdc = graphics.GetHdc();
                    // 使用 PrintWindow API，它比 BitBlt 更能應對複雜的 UI 視窗
                    NativeMethods.PrintWindow(hWnd, hdc, 0);
                    graphics.ReleaseHdc(hdc);
                }
                return bmp;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}