#region
// * DataWriter Class
//     This class provides the function to save data of streams (color, depth, body) to file.
//
// * License
//    Copyright (c) 2016 Tsukasa SUGIURA
//    This source code is licensed under the MIT license.
//
// * How to setup Microsoft.Kinect.KinectStudio
//     1. Add This Item to a Project.
//     2. Add Reference "System.Windows.Forms" to a Project.
//     5. Add "using Microsoft.Kinect.DataWriter;" to Your Source Code.
#endregion

using System;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Kinect;

namespace Microsoft.Kinect.DataWriter
{
    /// <summary>
    /// This class provides the function to save data of streams (color, depth, body) to file.
    /// </summary>
    public class DataWriter
    {
        private StreamWriter csv;
        private string directory;
        private DateTime time;
        private bool isSave;
        private bool isColor;
        private bool isDepth;
        private bool isBody;

        private byte[] colorBuffer;
        private ushort[] depthBuffer;
        private Body[] bodies;

        private WriteableBitmap colorBitmap;
        private Int32Rect colorRect;
        private int colorStride;

        private WriteableBitmap depthBitmap;
        private Int32Rect depthRect;
        private int depthStride;

        /// <summary>
        /// Constructor
        /// </summary>
        public DataWriter()
        {
            csv = null;
            directory = string.Empty;
            time = DateTime.Now;
            isSave = false;
            isColor = false;
            isDepth = false;
            isBody = false;
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~DataWriter()
        {
            if (csv != null)
            {
                csv.Dispose();
                csv = null;
            }
        }

        /// <summary>
        /// Start
        /// </summary>
        /// <param name="path">Save Data directory path.</param>
        /// <param name="isColor">Save Color frame.</param>
        /// <param name="isDepth">Save Depth frame.</param>
        /// <param name="isBody">Save Body frame.</param>
        public void Start(string path, bool isColor = false, bool isDepth = false, bool isBody = true)
        {
            if (isSave)
            {
                return;
            }

            if (!isColor && !isDepth && !isBody)
            {
                throw new ArgumentException("Not Specified Save Data");
            }

            if (!Directory.Exists(path))
            {
                throw new ArgumentException("Invalid Path");
            }

            directory = path;
            isSave = true;

            if (isColor)
            {
                FrameDescription description = KinectSensor.GetDefault().ColorFrameSource.CreateFrameDescription(ColorImageFormat.Bgra);
                colorBuffer = new byte[description.LengthInPixels * description.BytesPerPixel];
                colorRect = new Int32Rect(0, 0, description.Width, description.Height);
                colorStride = description.Width * (int)description.BytesPerPixel;
                colorBitmap = new WriteableBitmap(description.Width, description.Height, 96, 96, PixelFormats.Bgra32, null);

                this.isColor = isColor;
            }

            if (isDepth)
            {
                FrameDescription description = KinectSensor.GetDefault().DepthFrameSource.FrameDescription;
                depthBuffer = new ushort[description.LengthInPixels];
                depthRect = new Int32Rect(0, 0, description.Width, description.Height);
                depthStride = description.Width * (int)description.BytesPerPixel;
                depthBitmap = new WriteableBitmap(description.Width, description.Height, 96, 96, PixelFormats.Gray16, null);

                this.isDepth = isDepth;
            }

            if (isBody)
            {
                bodies = new Body[KinectSensor.GetDefault().BodyFrameSource.BodyCount];

                string filename = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.CurrentUICulture.DateTimeFormat);
                string extension = ".csv";

                csv = new StreamWriter(Path.Combine(directory, filename + extension), false, System.Text.Encoding.ASCII);

                AddLabel();

                this.isBody = isBody;
            }
        }

        /// <summary>
        /// Start with Save Foloder Choose Dialog
        /// </summary>
        /// <param name="isColor">Save Color frame.</param>
        /// <param name="isDepth">Save Depth frame.</param>
        /// <param name="isBody">Save Body frame.</param>
        public void Start(bool isColor = false, bool isDepth = false, bool isBody = true)
        {
            if (isSave)
            {
                return;
            }

            if (!isColor && !isDepth && !isBody)
            {
                throw new ArgumentException("Not Specified Save Data");
            }

            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Choose Save Folder";
                dialog.RootFolder = Environment.SpecialFolder.Desktop;
                dialog.ShowNewFolderButton = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    directory = dialog.SelectedPath;
                }
                else
                {
                    return;
                }
            }

            Start(directory, isColor, isDepth, isBody);
        }

        /// <summary>
        /// Add Label to CSV
        /// </summary>
        private void AddLabel()
        {
            if (csv == null)
            {
                return;
            }

            csv.Write(",");
            foreach (string joint in typeof(JointType).GetEnumNames())
            {
                csv.Write($",{joint},,");
            }
            csv.Write("\n");

            csv.Write("Retrieved Time,Tracking ID");
            for (int i = 0; i < typeof(JointType).GetEnumNames().Length; i++)
            {
                csv.Write(",x,y,z");
            }
            csv.Write("\n");
        }

        /// <summary>
        /// Stop
        /// </summary>
        public void Stop()
        {
            if (csv != null)
            {
                csv.Dispose();
                csv = null;
            }

            isSave = false;
        }

        /// <summary>
        /// Write
        /// </summary>
        /// <param name="multiFrame">MultiSourceFrame retrieved from Kinect.</param>
        public void Write(MultiSourceFrame multiFrame)
        {
            if (isSave)
            {
                time = System.DateTime.Now;

                if (isColor)
                {
                    WriteColor(multiFrame);
                }

                if (isDepth)
                {
                    WriteDepth(multiFrame);
                }

                if (isBody)
                {
                    WriteBody(multiFrame);
                }
            }
        }

        /// <summary>
        /// Write Color Frame
        /// </summary>
        /// <param name="multiFrame">MultiSourceFrame retrieved from Kinect.</param>
        private void WriteColor(MultiSourceFrame multiFrame)
        {
            if (multiFrame == null)
            {
                return;
            }

            using (ColorFrame colorFrame = multiFrame.ColorFrameReference.AcquireFrame())
            {
                if (colorFrame == null)
                {
                    return;
                }

                colorFrame.CopyConvertedFrameDataToArray(colorBuffer, ColorImageFormat.Bgra);
                colorBitmap.WritePixels(colorRect, colorBuffer, colorStride, 0);

                string filename = Regex.Replace(Regex.Replace(colorFrame.RelativeTime.ToString(), "[:]", ""), "[.]", "_");
                //string filename = time.ToString("yyyyMMdd_HHmmssfff", CultureInfo.CurrentUICulture.DateTimeFormat);
                string extension = ".jpg";

                using (FileStream stream = new FileStream(Path.Combine(directory, filename + extension), FileMode.Create, FileAccess.Write))
                {
                    JpegBitmapEncoder encoder = new JpegBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(colorBitmap));
                    encoder.Save(stream);
                }
            }
        }

        /// <summary>
        /// Write Depth Frame
        /// </summary>
        /// <param name="multiFrame">MultiSourceFrame retrieved from Kinect.</param>
        private void WriteDepth(MultiSourceFrame multiFrame)
        {
            if (multiFrame == null)
            {
                return;
            }

            using (DepthFrame depthFrame = multiFrame.DepthFrameReference.AcquireFrame())
            {
                if (depthFrame == null)
                {
                    return;
                }

                depthFrame.CopyFrameDataToArray(depthBuffer);
                depthBitmap.WritePixels(depthRect, depthBuffer, depthStride, 0);

                string filename = Regex.Replace(Regex.Replace(depthFrame.RelativeTime.ToString(), "[:]", ""), "[.]", "_");
                //string filename = time.ToString("yyyyMMdd_HHmmssfff", CultureInfo.CurrentUICulture.DateTimeFormat);
                string extension = ".png";

                using (FileStream stream = new FileStream(Path.Combine(directory, filename + extension), FileMode.Create, FileAccess.Write))
                {
                    PngBitmapEncoder encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(depthBitmap));
                    encoder.Save(stream);
                }
            }
        }

        /// <summary>
        /// Write Body Frame
        /// </summary>
        /// <param name="multiFrame">MultiSourceFrame retrieved from Kinect.</param>
        private void WriteBody(MultiSourceFrame multiFrame)
        {
            if (multiFrame == null)
            {
                return;
            }

            using (BodyFrame bodyFrame = multiFrame.BodyFrameReference.AcquireFrame())
            {
                if (bodyFrame == null)
                {
                    return;
                }

                bodyFrame.GetAndRefreshBodyData(bodies);

                string time = bodyFrame.RelativeTime.ToString();
                //string time = this.time.ToString("yyyy/MM/dd HH:mm:ss.fff", CultureInfo.CurrentUICulture.DateTimeFormat);

                foreach (Body body in bodies.Where(body => body != null))
                {
                    if (!body.IsTracked)
                    {
                        continue;
                    }

                    csv.Write($"{time},{body.TrackingId}");
                    foreach (var joint in body.Joints)
                    {
                        if (joint.Value.TrackingState == TrackingState.Tracked)
                        {
                            csv.Write($",{joint.Value.Position.X}");
                            csv.Write($",{joint.Value.Position.Y}");
                            csv.Write($",{joint.Value.Position.Z}");
                        }
                        else
                        {
                            csv.Write(",,,");
                        }
                    }
                    csv.Write("\n");
                }
            }
        }

    }
}
