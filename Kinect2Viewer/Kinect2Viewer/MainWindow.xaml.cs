using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Controls;
using Microsoft.Win32;
using Microsoft.Kinect;
using Microsoft.Kinect.KinectStudio;

namespace Kinect2Viewer
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        // Kinect
        KinectSensor kinect;
        CoordinateMapper coordinateMapper;

        MultiSourceFrameReader multiFrameReader;

        FrameDescription colorFrameDescription;
        FrameDescription depthFrameDescription;
        FrameDescription infraredFrameDescription;

        ColorImageFormat colorFormat = ColorImageFormat.Bgra;

        // Kinect Studio
        KinectStudio studio;

        // WPF
        WriteableBitmap colorBitmap;
        byte[] colorBuffer;
        Int32Rect colorRect;
        int colorStride;

        WriteableBitmap depthBitmap;
        ushort[] depthBuffer;
        byte[] buffer;
        Int32Rect depthRect;
        int depthStride;

        WriteableBitmap infraredBitmap;
        ushort[] infraredBuffer;
        Int32Rect infraredRect;
        int infraredStride;

        Body[] bodies;
        readonly Brush[] brushes = {
                        Brushes.Blue,
                        Brushes.Green,
                        Brushes.Red,
                        Brushes.Cyan,
                        Brushes.Yellow,
                        Brushes.Magenta
                    };

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Sensor
                kinect = KinectSensor.GetDefault();
                kinect.Open();

                // Coordinate Mapper
                coordinateMapper = kinect.CoordinateMapper;

                // Multi Reader
                FrameSourceTypes types = FrameSourceTypes.Color
                                       | FrameSourceTypes.Depth
                                       | FrameSourceTypes.Infrared
                                       | FrameSourceTypes.Body;
                multiFrameReader = kinect.OpenMultiSourceFrameReader(types);
                multiFrameReader.MultiSourceFrameArrived += MultiFrameArrived;

                // Description
                colorFrameDescription = kinect.ColorFrameSource.CreateFrameDescription(colorFormat);
                colorRect = new Int32Rect(0, 0, colorFrameDescription.Width, colorFrameDescription.Height);
                colorStride = colorFrameDescription.Width * (int)colorFrameDescription.BytesPerPixel;
                colorBuffer = new byte[colorFrameDescription.LengthInPixels * colorFrameDescription.BytesPerPixel];

                depthFrameDescription = kinect.DepthFrameSource.FrameDescription;
                depthRect = new Int32Rect(0, 0, depthFrameDescription.Width, depthFrameDescription.Height);
                depthStride = depthFrameDescription.Width;
                depthBuffer = new ushort[depthFrameDescription.LengthInPixels];
                buffer = new byte[depthFrameDescription.LengthInPixels];

                infraredFrameDescription = kinect.InfraredFrameSource.FrameDescription;
                infraredRect = new Int32Rect(0, 0, infraredFrameDescription.Width, infraredFrameDescription.Height);
                infraredStride = infraredFrameDescription.Width * (int)infraredFrameDescription.BytesPerPixel;
                infraredBuffer = new ushort[infraredFrameDescription.LengthInPixels];

                bodies = new Body[kinect.BodyFrameSource.BodyCount];

                // Bitmap
                colorBitmap = new WriteableBitmap(colorFrameDescription.Width, colorFrameDescription.Height, 96, 96, PixelFormats.Bgra32, null);
                depthBitmap = new WriteableBitmap(depthFrameDescription.Width, depthFrameDescription.Height, 96, 96, PixelFormats.Gray8, null);
                infraredBitmap = new WriteableBitmap(infraredFrameDescription.Width, infraredFrameDescription.Height, 96, 96, PixelFormats.Gray16, null);

                Color.Source = colorBitmap;
                Depth.Source = depthBitmap;
                Infrared.Source = infraredBitmap;

                // Kinect Studio
                studio = new KinectStudio();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Close();
            }
        }

        private void MultiFrameArrived(object sender, MultiSourceFrameArrivedEventArgs e)
        {
            MultiSourceFrame multiFrame = e.FrameReference.AcquireFrame();
            if (multiFrame == null)
            {
                return;
            }

            // Update Frame
            UpdateColorFrame(multiFrame);
            UpdateDepthFrame(multiFrame);
            UpdateInfraredFrame(multiFrame);
            UpdateBodyFrame(multiFrame);

            // Draw Frame
            DrawColorFrame();
            DrawDepthFrame();
            DrawInfraredFrame();
            DrawBodyFrame();
        }

        private void UpdateColorFrame(MultiSourceFrame multiFrame)
        {
            // Frame
            using (ColorFrame colorFrame = multiFrame.ColorFrameReference.AcquireFrame())
            {
                if (colorFrame == null)
                {
                    return;
                }

                colorFrame.CopyConvertedFrameDataToArray(colorBuffer, ColorImageFormat.Bgra);
            }
        }

        private void UpdateDepthFrame(MultiSourceFrame multiFrame)
        {
            // Frame
            using (DepthFrame depthFrame = multiFrame.DepthFrameReference.AcquireFrame())
            {
                if (depthFrame == null)
                {
                    return;
                }

                depthFrame.CopyFrameDataToArray(depthBuffer);
            }
        }

        private void UpdateInfraredFrame(MultiSourceFrame multiFrame)
        {
            // Frame
            using (InfraredFrame infraredFrame = multiFrame.InfraredFrameReference.AcquireFrame())
            {
                if (infraredFrame == null)
                {
                    return;
                }

                infraredFrame.CopyFrameDataToArray(infraredBuffer);
            }
        }

        private void UpdateBodyFrame(MultiSourceFrame multiFrame)
        {
            // Frame
            using (BodyFrame bodyFrame = multiFrame.BodyFrameReference.AcquireFrame())
            {
                if (bodyFrame == null)
                {
                    return;
                }

                bodyFrame.GetAndRefreshBodyData(bodies);
            }
        }

        private void DrawColorFrame()
        {
            colorBitmap.WritePixels(colorRect, colorBuffer, colorStride, 0);
        }

        private void DrawDepthFrame()
        {
            float alpha = -(byte.MaxValue / 8000.0f);
            float beta = 255.0f;
            for (int i = 0; i < depthBuffer.Length; i++)
            {
                buffer[i] = (byte)(alpha * depthBuffer[i] + beta);
            }

            depthBitmap.WritePixels(depthRect, buffer, depthStride, 0);
        }

        private void DrawInfraredFrame()
        {
            infraredBitmap.WritePixels(infraredRect, infraredBuffer, infraredStride, 0);
        }

        private void DrawBodyFrame()
        {
            Body.Children.Clear();

            for (int i = 0; i < bodies.Length; i++)
            {
                Body body = bodies[i];
                if (body != null)
                {
                    if (!body.IsTracked)
                    {
                        continue;
                    }

                    foreach (var joint in body.Joints)
                    {
                        if (joint.Value.TrackingState != TrackingState.NotTracked)
                        {
                            DrawEllipse(joint.Value, 10, brushes[i]);
                        }
                    }
                }
            }
        }

        private void DrawEllipse(Joint joint, int radius, Brush brush, double alpha = 1.0)
        {
            Ellipse ellipse = new Ellipse()
            {
                Width = radius,
                Height = radius,
                Fill = brush,
                Opacity = alpha,
            };

            ColorSpacePoint point = coordinateMapper.MapCameraPointToColorSpace(joint.Position);
            if ((0 > point.X) || (point.X >= colorBitmap.Width) || (0 > point.Y) || (point.Y >= colorBitmap.Height))
            {
                return;
            }

            double coef = Color.Width / colorBitmap.Width;
            Canvas.SetLeft(ellipse, (point.X - (radius / 2)) * coef);
            Canvas.SetTop(ellipse, (point.Y - (radius / 2)) * coef);

            Body.Children.Add(ellipse);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (studio != null)
            {
                studio.Stop();
                studio = null;
            }

            if (multiFrameReader != null)
            {
                multiFrameReader.Dispose();
                multiFrameReader = null;
            }

            if (kinect != null)
            {
                kinect.Close();
                kinect = null;
            }
        }

        private void Open_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Open Clip File";
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Kinect Studio\\Repository";
            dialog.Filter = "Event File|*.xef";

            if (dialog.ShowDialog() == true)
            {
                studio.Clip(dialog.FileName);
            }
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ClipPlay_Click(object sender, RoutedEventArgs e)
        {
            studio.Play();
        }

        private void ClipPause_Click(object sender, RoutedEventArgs e)
        {
            studio.Pause();
        }

        private void ClipStop_Click(object sender, RoutedEventArgs e)
        {
            studio.Stop();
        }

        private void DrawColor_Checked(object sender, RoutedEventArgs e)
        {
            if (Color != null)
            {
                Color.Visibility = Visibility.Visible;
            }
        }

        private void DrawBody_Checked(object sender, RoutedEventArgs e)
        {
            if (Body != null)
            {
                Body.Visibility = Visibility.Visible;
            }
        }

        private void DrawDepth_Checked(object sender, RoutedEventArgs e)
        {
            if (Depth != null && DrawInfrared != null)
            {
                Depth.Visibility = Visibility.Visible;
                DrawInfrared.IsChecked = false;
            }
        }

        private void DrawInfrared_Checked(object sender, RoutedEventArgs e)
        {
            if (Infrared != null && DrawDepth != null)
            {
                Infrared.Visibility = Visibility.Visible;
                DrawDepth.IsChecked = false;
            }
        }

        private void DrawColor_Unchecked(object sender, RoutedEventArgs e)
        {
            if (Color != null)
            {
                Color.Visibility = Visibility.Hidden;
            }
        }

        private void DrawBody_Unchecked(object sender, RoutedEventArgs e)
        {
            if (Body != null)
            {
                Body.Visibility = Visibility.Hidden;
            }
        }

        private void DrawDepth_Unchecked(object sender, RoutedEventArgs e)
        {
            if (Depth != null)
            {
                Depth.Visibility = Visibility.Hidden;
            }
        }

        private void DrawInfrared_Unchecked(object sender, RoutedEventArgs e)
        {
            if (Infrared != null)
            {
                Infrared.Visibility = Visibility.Hidden;
            }
        }
    }
}

