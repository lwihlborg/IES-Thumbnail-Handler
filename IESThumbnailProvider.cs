using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;

namespace IESThumbnailHandler
{
    /// <summary>
    /// Windows Shell Thumbnail Provider for IES Photometry Files
    /// Displays polar plots of vertical and horizontal candela distributions
    /// </summary>
    [ComVisible(true)]
    [Guid("7B3FC2A1-E8D4-4F5B-9A2C-8E1D3F6B5C4A")]
    [ClassInterface(ClassInterfaceType.None)]
    public class IESThumbnailProvider : IThumbnailProvider, IInitializeWithStream
    {
        private IStream _stream;

        #region IInitializeWithStream
        public int Initialize(IStream pstream, uint grfMode)
        {
            _stream = pstream;
            return 0; // S_OK
        }
        #endregion

        #region IThumbnailProvider
        public int GetThumbnail(uint cx, out IntPtr phbmp, out WTS_ALPHATYPE pdwAlpha)
        {
            phbmp = IntPtr.Zero;
            pdwAlpha = WTS_ALPHATYPE.WTSAT_ARGB;

            try
            {
                // Read stream content
                string content = ReadStreamContent();
                if (string.IsNullOrEmpty(content))
                    return -1; // E_FAIL

                // Parse IES file
                IESData ies = new IESData(content);

                // Generate thumbnail
                int size = (int)Math.Min(cx, 256);
                using (Bitmap bitmap = GenerateThumbnail(ies, size))
                {
                    phbmp = bitmap.GetHbitmap(Color.FromArgb(0, 0, 0, 0));
                }

                return 0; // S_OK
            }
            catch (OutOfMemoryException)
            {
                throw;
            }
            catch
            {
                return -1; // E_FAIL
            }
        }
        #endregion

        private string ReadStreamContent()
        {
            if (_stream == null) return null;

            const int maxFileSize = 10 * 1024 * 1024;
            const int bufferSize = 4096;
            byte[] buffer = new byte[bufferSize];
            using (MemoryStream ms = new MemoryStream())
            {
                int bytesRead;
                do
                {
                    _stream.Read(buffer, bufferSize, out bytesRead);
                    if (bytesRead > 0)
                        ms.Write(buffer, 0, bytesRead);
                    if (ms.Length > maxFileSize)
                        return null;
                } while (bytesRead > 0);

                return System.Text.Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        private Bitmap GenerateThumbnail(IESData ies, int size)
        {
            Bitmap bitmap = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.FromArgb(255, 255, 255, 255)); // White background

                // Colors - pure blue for vertical, pure red for horizontal
                Color vertColor = Color.FromArgb(255, 0, 0, 255);      // #0000FF
                Color horizColor = Color.FromArgb(255, 255, 0, 0);     // #FF0000
                Color gridColor = Color.FromArgb(180, 180, 180, 180);  // Light gray

                // Single centered plot
                int padding = (int)(size * 0.08);
                int radius = (size / 2) - padding - 2;
                Point center = new Point(size / 2, size / 2);

                // Draw grid
                using (Pen gridPen = new Pen(gridColor, 1))
                {
                    // Concentric circles
                    float[] fractions = new float[] { 0.25f, 0.5f, 0.75f, 1.0f };
                    foreach (float frac in fractions)
                    {
                        int r = (int)(radius * frac);
                        g.DrawEllipse(gridPen, center.X - r, center.Y - r, r * 2, r * 2);
                    }
                    
                    // Radial lines every 30 degrees (12 lines)
                    for (int deg = 0; deg < 360; deg += 30)
                    {
                        double theta = deg * Math.PI / 180;
                        int x2 = (int)(center.X + radius * Math.Sin(theta));
                        int y2 = (int)(center.Y - radius * Math.Cos(theta));
                        g.DrawLine(gridPen, center.X, center.Y, x2, y2);
                    }
                }

                // Get distributions
                List<double> vertAngles;
                List<double> vertCandela;
                ies.GetVerticalDistribution(out vertAngles, out vertCandela);
                
                List<double> horizAngles;
                List<double> horizCandela;
                ies.GetHorizontalDistribution(out horizAngles, out horizCandela);
                ExpandHorizontalSymmetry(ref horizAngles, ref horizCandela);
                
                double maxCd = 1;
                if (vertCandela.Count > 0) maxCd = Math.Max(maxCd, vertCandela.Max());
                if (horizCandela.Count > 0) maxCd = Math.Max(maxCd, horizCandela.Max());

                // Line thickness scales with thumbnail size
                float lineWidth = Math.Max(2.0f, size / 80.0f);

                // Draw horizontal distribution first (red, background)
                DrawHorizontalDistribution(g, horizAngles, horizCandela, center, radius, maxCd, horizColor, lineWidth);

                // Draw vertical distribution on top (blue, foreground)
                DrawVerticalDistribution(g, vertAngles, vertCandela, center, radius, maxCd, vertColor, lineWidth);

                // Info text (lumens / watts) — only on Large (96) and Extra Large (256) thumbnails
                if (size >= 96)
                    DrawInfoText(g, ies, size);
            }

            return bitmap;
        }

        private void DrawInfoText(Graphics g, IESData ies, int size)
        {
            double lumens = ies.TotalLumens;
            double watts = ies.InputWatts;

            List<string> parts = new List<string>();
            if (lumens > 0.5)
                parts.Add(lumens.ToString("#,##0", System.Globalization.CultureInfo.CurrentCulture) + " lm");
            if (watts > 0.05)
                parts.Add(watts.ToString("0.#", System.Globalization.CultureInfo.CurrentCulture) + "W");
            if (parts.Count == 0) return;

            string text = string.Join("   ", parts.ToArray());

            const float fontSize = 14f;
            const float margin = 4f;
            System.Drawing.Text.TextRenderingHint prevHint = g.TextRenderingHint;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            try
            {
                using (Font font = new Font("Segoe UI", fontSize, FontStyle.Regular, GraphicsUnit.Pixel))
                using (Brush brush = new SolidBrush(Color.FromArgb(230, 40, 40, 40)))
                {
                    StringFormat sf = StringFormat.GenericTypographic;
                    SizeF textSize = g.MeasureString(text, font, int.MaxValue, sf);
                    float x = size - textSize.Width - margin;
                    float y = margin;
                    g.DrawString(text, font, brush, x, y, sf);
                }
            }
            finally
            {
                g.TextRenderingHint = prevHint;
            }
        }

        private void DrawVerticalDistribution(Graphics g, List<double> angles, List<double> candela,
            Point center, int radius, double maxCd, Color plotColor, float lineWidth)
        {
            if (candela.Count == 0 || maxCd == 0) return;

            List<PointF> points = new List<PointF>();

            // Right side points (0° at bottom, going clockwise)
            for (int i = 0; i < angles.Count; i++)
            {
                double r = (candela[i] / maxCd) * radius;
                double theta = angles[i] * Math.PI / 180;
                float x = (float)(center.X + r * Math.Sin(theta));
                float y = (float)(center.Y + r * Math.Cos(theta));
                points.Add(new PointF(x, y));
            }

            // Mirror for left side
            for (int i = angles.Count - 1; i >= 0; i--)
            {
                double r = (candela[i] / maxCd) * radius;
                double theta = -angles[i] * Math.PI / 180;
                float x = (float)(center.X + r * Math.Sin(theta));
                float y = (float)(center.Y + r * Math.Cos(theta));
                points.Add(new PointF(x, y));
            }

            if (points.Count >= 3)
            {
                using (Pen pen = new Pen(plotColor, lineWidth))
                {
                    g.DrawPolygon(pen, points.ToArray());
                }
            }
        }

        private void DrawHorizontalDistribution(Graphics g, List<double> angles, List<double> candela,
            Point center, int radius, double maxCd, Color plotColor, float lineWidth)
        {
            if (candela.Count == 0 || maxCd == 0) return;

            List<PointF> points = new List<PointF>();

            for (int i = 0; i < angles.Count; i++)
            {
                double r = (candela[i] / maxCd) * radius;
                double theta = (angles[i] - 90) * Math.PI / 180;
                float x = (float)(center.X + r * Math.Cos(theta));
                float y = (float)(center.Y + r * Math.Sin(theta));
                points.Add(new PointF(x, y));
            }

            // Close the polygon
            if (points.Count > 0)
                points.Add(points[0]);

            if (points.Count >= 3)
            {
                using (Pen pen = new Pen(plotColor, lineWidth))
                {
                    g.DrawPolygon(pen, points.ToArray());
                }
            }
        }

        private void ExpandHorizontalSymmetry(ref List<double> angles, ref List<double> candela)
        {
            if (angles.Count == 0)
            {
                angles = new List<double> { 0, 90, 180, 270 };
                candela = new List<double> { 0, 0, 0, 0 };
                return;
            }

            double maxH = angles.Max();
            
            // Use dictionary to handle duplicates and sorting
            SortedDictionary<double, double> angleCD = new SortedDictionary<double, double>();
            for (int i = 0; i < angles.Count; i++)
                angleCD[angles[i]] = candela[i];

            if (maxH <= 0)
            {
                // Single point - circular distribution
                double c = candela[0];
                angleCD[90] = c;
                angleCD[180] = c;
                angleCD[270] = c;
            }
            else if (maxH <= 90)
            {
                // Quadrant symmetry - mirror to full 360
                List<KeyValuePair<double, double>> items = new List<KeyValuePair<double, double>>(angleCD);
                foreach (KeyValuePair<double, double> kvp in items)
                {
                    if (kvp.Key < 90) angleCD[180 - kvp.Key] = kvp.Value;
                }
                items = new List<KeyValuePair<double, double>>(angleCD);
                foreach (KeyValuePair<double, double> kvp in items)
                {
                    if (kvp.Key > 0 && kvp.Key < 180) angleCD[360 - kvp.Key] = kvp.Value;
                }
            }
            else if (maxH <= 180)
            {
                // Bilateral symmetry
                List<KeyValuePair<double, double>> items = new List<KeyValuePair<double, double>>(angleCD);
                foreach (KeyValuePair<double, double> kvp in items)
                {
                    if (kvp.Key > 0 && kvp.Key < 360) angleCD[360 - kvp.Key] = kvp.Value;
                }
            }

            angles = new List<double>(angleCD.Keys);
            candela = new List<double>(angleCD.Values);
        }

        #region COM Registration
        [ComRegisterFunction]
        public static void Register(Type t)
        {
            try
            {
                string guid = t.GUID.ToString("B");
                string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;

                // Register CLSID per-user
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\Classes\\CLSID\\" + guid))
                {
                    if (key != null) key.SetValue(null, "IES Photometry Thumbnail Handler");
                }
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\Classes\\CLSID\\" + guid + "\\InprocServer32"))
                {
                    if (key != null)
                    {
                        key.SetValue(null, dllPath);
                        key.SetValue("ThreadingModel", "Apartment");
                    }
                }

                // Register for .ies extension
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\Classes\\.ies\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}"))
                {
                    if (key != null) key.SetValue(null, guid);
                }

                // Register for .IES extension (uppercase)
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\Classes\\.IES\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}"))
                {
                    if (key != null) key.SetValue(null, guid);
                }

                // Mark as approved (per-user)
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(
                    "Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved"))
                {
                    if (key != null) key.SetValue(guid, "IES Photometry Thumbnail Handler");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Registration failed: " + ex.Message);
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            try
            {
                string guid = t.GUID.ToString("B");

                Registry.CurrentUser.DeleteSubKeyTree("Software\\Classes\\CLSID\\" + guid, false);
                Registry.CurrentUser.DeleteSubKeyTree("Software\\Classes\\.ies\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}", false);
                Registry.CurrentUser.DeleteSubKeyTree("Software\\Classes\\.IES\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}", false);

                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(
                    "Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved", true))
                {
                    if (key != null) key.DeleteValue(guid, false);
                }
            }
            catch (OutOfMemoryException)
            {
                throw;
            }
            catch { }
        }
        #endregion
    }

    /// <summary>
    /// Lightweight IES file parser
    /// </summary>
    internal class IESData
    {
        public List<double> VerticalAngles { get; private set; }
        public List<double> HorizontalAngles { get; private set; }
        public List<List<double>> CandelaValues { get; private set; }
        public double NumLamps { get; private set; }
        public double LumensPerLamp { get; private set; }
        public double InputWatts { get; private set; }
        public bool IsAbsolute { get { return LumensPerLamp <= 0; } }
        public double TotalLumens
        {
            get
            {
                if (!IsAbsolute && NumLamps > 0)
                    return NumLamps * LumensPerLamp;
                return CalculateLumens();
            }
        }
        private double _candelaMultiplier = 1.0;

        public IESData(string content)
        {
            VerticalAngles = new List<double>();
            HorizontalAngles = new List<double>();
            CandelaValues = new List<List<double>>();
            Parse(content);
        }

        private void Parse(string content)
        {
            string[] lines = content.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> lineList = new List<string>();
            foreach (string line in lines)
            {
                lineList.Add(line.Trim());
            }

            // Find TILT line
            int tiltIndex = -1;
            for (int i = 0; i < lineList.Count; i++)
            {
                if (lineList[i].ToUpper().StartsWith("TILT="))
                {
                    tiltIndex = i;
                    break;
                }
            }
            if (tiltIndex < 0) throw new FormatException("Invalid IES file");

            // Collect numeric data
            List<double> numericData = new List<double>();
            for (int i = tiltIndex + 1; i < lineList.Count; i++)
            {
                string[] values = Regex.Split(lineList[i], "[\\s,]+");
                foreach (string v in values)
                {
                    double num;
                    if (double.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out num))
                        numericData.Add(num);
                }
            }

            if (numericData.Count < 13) throw new FormatException("Insufficient data");

            // Parse header
            int idx = 0;
            NumLamps = numericData[idx++];
            LumensPerLamp = numericData[idx++];
            _candelaMultiplier = numericData[idx++];
            int numVert = (int)numericData[idx++];
            int numHoriz = (int)numericData[idx++];

            if (numVert < 0 || numVert > 10000 || numHoriz < 0 || numHoriz > 10000)
                throw new FormatException("Invalid IES data dimensions");
            if (idx + 8 > numericData.Count)
                throw new FormatException("Invalid IES data dimensions");

            idx += 7; // photometric_type, units_type, width, length, height, ballast_factor, future_use
            InputWatts = numericData[idx++];

            if (numericData.Count < idx + numVert + numHoriz + numVert * numHoriz)
                throw new FormatException("Invalid IES data dimensions");

            // Parse angles
            for (int i = 0; i < numVert; i++)
            {
                VerticalAngles.Add(numericData[idx + i]);
            }
            idx += numVert;

            for (int i = 0; i < numHoriz; i++)
            {
                HorizontalAngles.Add(numericData[idx + i]);
            }
            idx += numHoriz;

            // Parse candela values
            for (int h = 0; h < numHoriz; h++)
            {
                List<double> candela = new List<double>();
                for (int v = 0; v < numVert; v++)
                {
                    candela.Add(numericData[idx + v] * _candelaMultiplier);
                }
                CandelaValues.Add(candela);
                idx += numVert;
            }
        }

        // Total flux by zonal integration of the candela distribution.
        // Trapezoidal in the horizontal range gives the average intensity per
        // vertical zone; symmetry of the IES horizontal range makes that average
        // equivalent to a full 360° average.
        private double CalculateLumens()
        {
            int numVert = VerticalAngles.Count;
            int numHoriz = HorizontalAngles.Count;
            if (numVert < 2 || numHoriz < 1 || CandelaValues.Count < numHoriz)
                return 0;

            double[] avgI = new double[numVert];
            double hRange = numHoriz > 1 ? (HorizontalAngles[numHoriz - 1] - HorizontalAngles[0]) : 0;

            for (int v = 0; v < numVert; v++)
            {
                if (numHoriz == 1 || hRange <= 0)
                {
                    avgI[v] = CandelaValues[0].Count > v ? CandelaValues[0][v] : 0;
                }
                else
                {
                    double sum = 0;
                    for (int h = 0; h < numHoriz - 1; h++)
                    {
                        double i1 = CandelaValues[h].Count > v ? CandelaValues[h][v] : 0;
                        double i2 = CandelaValues[h + 1].Count > v ? CandelaValues[h + 1][v] : 0;
                        sum += (i1 + i2) / 2.0 * (HorizontalAngles[h + 1] - HorizontalAngles[h]);
                    }
                    avgI[v] = sum / hRange;
                }
            }

            double total = 0;
            for (int v = 0; v < numVert - 1; v++)
            {
                double t1 = VerticalAngles[v] * Math.PI / 180.0;
                double t2 = VerticalAngles[v + 1] * Math.PI / 180.0;
                double zonalOmega = 2 * Math.PI * Math.Abs(Math.Cos(t1) - Math.Cos(t2));
                total += (avgI[v] + avgI[v + 1]) / 2.0 * zonalOmega;
            }
            return total;
        }

        public void GetVerticalDistribution(out List<double> angles, out List<double> candela)
        {
            angles = new List<double>(VerticalAngles);
            if (CandelaValues.Count > 0)
            {
                candela = new List<double>(CandelaValues[0]);
            }
            else
            {
                candela = new List<double>();
            }
        }

        public void GetHorizontalDistribution(out List<double> angles, out List<double> candela)
        {
            angles = new List<double>(HorizontalAngles);
            candela = new List<double>();
            
            if (VerticalAngles.Count == 0 || CandelaValues.Count == 0)
            {
                return;
            }

            // Find angle closest to 45 degrees
            int vertIdx = 0;
            double minDiff = double.MaxValue;
            for (int i = 0; i < VerticalAngles.Count; i++)
            {
                double diff = Math.Abs(VerticalAngles[i] - 45);
                if (diff < minDiff)
                {
                    minDiff = diff;
                    vertIdx = i;
                }
            }

            for (int h = 0; h < CandelaValues.Count; h++)
            {
                if (CandelaValues[h].Count > vertIdx)
                {
                    candela.Add(CandelaValues[h][vertIdx]);
                }
                else
                {
                    candela.Add(0);
                }
            }
        }
    }

    #region COM Interfaces
    [ComImport]
    [Guid("e357fccd-a995-4576-b01f-234630154e96")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IThumbnailProvider
    {
        [PreserveSig]
        int GetThumbnail(uint cx, out IntPtr phbmp, out WTS_ALPHATYPE pdwAlpha);
    }

    [ComImport]
    [Guid("b824b49d-22ac-4161-ac8a-9916e8fa3f7f")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IInitializeWithStream
    {
        [PreserveSig]
        int Initialize(IStream pstream, uint grfMode);
    }

    [ComImport]
    [Guid("0000000c-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IStream
    {
        void Read([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, int cb, out int pcbRead);
        void Write([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, int cb, out int pcbWritten);
        void Seek(long dlibMove, int dwOrigin, out long plibNewPosition);
        void SetSize(long libNewSize);
        void CopyTo(IStream pstm, long cb, out long pcbRead, out long pcbWritten);
        void Commit(int grfCommitFlags);
        void Revert();
        void LockRegion(long libOffset, long cb, int dwLockType);
        void UnlockRegion(long libOffset, long cb, int dwLockType);
        void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, int grfStatFlag);
        void Clone(out IStream ppstm);
    }

    public enum WTS_ALPHATYPE
    {
        WTSAT_UNKNOWN = 0,
        WTSAT_RGB = 1,
        WTSAT_ARGB = 2
    }
    #endregion
}
