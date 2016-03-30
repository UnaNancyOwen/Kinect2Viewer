using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Controls;
using Microsoft.Kinect;

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

        ColorFrameReader colorFrameReader;
        FrameDescription colorFrameDescription;
        ColorImageFormat colorFormat = ColorImageFormat.Bgra;

        DepthFrameReader depthFrameReader;
        FrameDescription depthFrameDescription;

        InfraredFrameReader infraredFrameReader;
        FrameDescription infraredFrameDescription;

        BodyFrameReader bodyFrameReader;

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

                // Source
                ColorFrameSource colorFrameSource = kinect.ColorFrameSource;
                DepthFrameSource depthFrameSource = kinect.DepthFrameSource;
                InfraredFrameSource infraredFrameSource = kinect.InfraredFrameSource;
                BodyFrameSource bodyFrameSource = kinect.BodyFrameSource;

                // Reader
                colorFrameReader = colorFrameSource.OpenReader();
                colorFrameReader.FrameArrived += ColorFrameArrived;

                depthFrameReader = depthFrameSource.OpenReader();
                depthFrameReader.FrameArrived += DepthFrameArrived;

                infraredFrameReader = infraredFrameSource.OpenReader();
                infraredFrameReader.FrameArrived += InfraredFrameArrived;

                bodyFrameReader = bodyFrameSource.OpenReader();
                bodyFrameReader.FrameArrived += bodyFrameArrived;

                // Description
                colorFrameDescription = colorFrameSource.CreateFrameDescription(colorFormat);
                colorRect = new Int32Rect(0, 0, colorFrameDescription.Width, colorFrameDescription.Height);
                colorStride = colorFrameDescription.Width * (int)colorFrameDescription.BytesPerPixel;

                depthFrameDescription = depthFrameSource.FrameDescription;
                depthRect = new Int32Rect(0, 0, depthFrameDescription.Width, depthFrameDescription.Height);
                depthStride = depthFrameDescription.Width;

                infraredFrameDescription = infraredFrameSource.FrameDescription;
                infraredRect = new Int32Rect(0, 0, infraredFrameDescription.Width, infraredFrameDescription.Height);
                infraredStride = infraredFrameDescription.Width * (int)infraredFrameDescription.BytesPerPixel;

                // Buffer
                colorBuffer = new byte[colorFrameDescription.LengthInPixels * colorFrameDescription.BytesPerPixel];

                depthBuffer = new ushort[depthFrameDescription.LengthInPixels];
                buffer = new byte[depthFrameDescription.LengthInPixels];

                infraredBuffer = new ushort[infraredFrameDescription.LengthInPixels];

                bodies = new Body[bodyFrameSource.BodyCount];

                // Bitmap
                colorBitmap = new WriteableBitmap(colorFrameDescription.Width, colorFrameDescription.Height, 96, 96, PixelFormats.Bgra32, null);
                depthBitmap = new WriteableBitmap(depthFrameDescription.Width, depthFrameDescription.Height, 96, 96, PixelFormats.Gray8, null);
                infraredBitmap = new WriteableBitmap(infraredFrameDescription.Width, infraredFrameDescription.Height, 96, 96, PixelFormats.Gray16, null);

                Color.Source = colorBitmap;
                Depth.Source = depthBitmap;
                Infrared.Source = infraredBitmap;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Close();
            }
        }

        private void ColorFrameArrived(object sender, ColorFrameArrivedEventArgs e)
        {
            UpdateColorFrame(e);
            DrawColorFrame();
        }

        private void DepthFrameArrived(object sender, DepthFrameArrivedEventArgs e)
        {
            UpdateDepthFrame(e);
            DrawDepthFrame();
        }

        private void InfraredFrameArrived(object sender, InfraredFrameArrivedEventArgs e)
        {
            UpdateInfraredFrame(e);
            DrawInfraredFrame();
        }

        private void bodyFrameArrived(object sender, BodyFrameArrivedEventArgs e)
        {
            UpdateBodyFrame(e);
            DrawBodyFrame();
        }

        private void UpdateColorFrame(ColorFrameArrivedEventArgs e)
        {
            // Frame
            using (ColorFrame colorFrame = e.FrameReference.AcquireFrame())
            {
                if (colorFrame == null)
                {
                    return;
                }

                colorFrame.CopyConvertedFrameDataToArray(colorBuffer, ColorImageFormat.Bgra);
            }
        }

        private void UpdateDepthFrame(DepthFrameArrivedEventArgs e)
        {
            // Frame
            using (DepthFrame depthFrame = e.FrameReference.AcquireFrame())
            {
                if (depthFrame == null)
                {
                    return;
                }

                depthFrame.CopyFrameDataToArray(depthBuffer);
            }
        }

        private void UpdateInfraredFrame(InfraredFrameArrivedEventArgs e)
        {
            // Frame
            using (InfraredFrame infraredFrame = e.FrameReference.AcquireFrame())
            {
                if (infraredFrame == null)
                {
                    return;
                }

                infraredFrame.CopyFrameDataToArray(infraredBuffer);
            }
        }

        private void UpdateBodyFrame(BodyFrameArrivedEventArgs e)
        {
            // Frame
            using (BodyFrame bodyFrame = e.FrameReference.AcquireFrame())
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
            if (colorFrameReader != null)
            {
                colorFrameReader.Dispose();
                colorFrameReader = null;
            }

            if (depthFrameReader != null)
            {
                depthFrameReader.Dispose();
                depthFrameReader = null;
            }

            if (infraredFrameReader != null)
            {
                infraredFrameReader.Dispose();
                infraredFrameReader = null;
            }

            if (bodyFrameReader != null)
            {
                bodyFrameReader.Dispose();
                bodyFrameReader = null;
            }

            if (kinect != null)
            {
                kinect.Close();
                kinect = null;
            }
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

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
