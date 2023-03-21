using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Kinect;
using System.Runtime.InteropServices;

namespace PowerPointKinect
{
    public partial class MainWindow : Window
    {
        KinectSensor myKinect;

        WriteableBitmap bitmapImagenColor = null;
        byte[] bytesColor;

        Skeleton[] skeleton = null;

        bool isCirclesVisible = false;
        bool isDragOn = false;
        bool mouse = false;
        bool selectLayer = false;
        bool movementZoomUp = false;
        bool movementZoomDown = false;
        bool movementFrontActive= false;
        bool movementBackActive = false;

        int milliseconds = 2000;

        double originX, originY;
        double[] x = new double[3];
        double[] y = new double[3];

        SolidColorBrush brushActivo = new SolidColorBrush(Colors.Green);
        SolidColorBrush brushInactivo = new SolidColorBrush(Colors.Red);

        [DllImport("user32")]
        public static extern int SetCursorPos(int x, int y);

        private const int MOUSEEVENTF_MOVE = 0x01;

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;

        private const int MOUSEEVENTF_LEFTUP = 0x04;

        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;

        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        private const int MOUSEEVENTF_MIDDLEDOWN = 0x20;

        private const int MOUSEEVENTF_MIDDLEUP = 0x40;

        private const int MOUSEEVENTF_ABSOLUTE = 0x80;

        [DllImport("user32.dll",
            CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]

        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);


        public MainWindow()
        {
            InitializeComponent();

            this.Loaded += new RoutedEventHandler(Window_Loaded_1);

            this.KeyDown += new KeyEventHandler(MainWindow_KeyDown);

        }

        private void Window_Loaded_1(object sender, RoutedEventArgs e)
        {
            myKinect = KinectSensor.KinectSensors.FirstOrDefault();
            if (myKinect == null)
            {
                MessageBox.Show("Need a Kinect");
                Application.Current.Shutdown();
            }

            myKinect.Start();
            myKinect.ColorStream.Enable();
            myKinect.SkeletonStream.Enable();

            myKinect.ColorFrameReady += myKinect_ColorFrameReady;
            myKinect.SkeletonFrameReady += myKinect_SkeletonFrameReady;
            Application.Current.Exit += Current_Exit;
        }

        void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.C)
            {
                ToggleCircles();
            }
        }

        void HideCircles()
        {
            isCirclesVisible = false;
            ellipseHead.Visibility = System.Windows.Visibility.Collapsed;
            ellipseHandLeft.Visibility = System.Windows.Visibility.Collapsed;
            ellipseHandRight.Visibility = System.Windows.Visibility.Collapsed;
        }

        void ShowCircles()
        {
            isCirclesVisible = true;
            ellipseHead.Visibility = System.Windows.Visibility.Visible;
            ellipseHandLeft.Visibility = System.Windows.Visibility.Visible;
            ellipseHandRight.Visibility = System.Windows.Visibility.Visible;
        }
        void ToggleCircles()
        {
            if (isCirclesVisible)
                HideCircles();
            else
                ShowCircles();
        }

        void Current_Exit(object sender, ExitEventArgs e)
        {
            if (myKinect != null) {
                myKinect.Stop();
                myKinect = null;
            }
        }

        void myKinect_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            using (SkeletonFrame frame = e.OpenSkeletonFrame())
            {
                if (frame != null)
                {
                    skeleton = new Skeleton[frame.SkeletonArrayLength];
                    frame.CopySkeletonDataTo(skeleton);
                }
            }

            if (skeleton == null) return;

            Skeleton esqueletoCercano = skeleton.Where(s => s.TrackingState == SkeletonTrackingState.Tracked)
                                                 .OrderBy(s => s.Position.Z * Math.Abs(s.Position.X))
                                                 .FirstOrDefault();

            if (esqueletoCercano == null) return;

            var head = esqueletoCercano.Joints[JointType.Head];
            var handRight = esqueletoCercano.Joints[JointType.HandRight];
            var handLeft = esqueletoCercano.Joints[JointType.HandLeft];

            if (head.TrackingState == JointTrackingState.NotTracked ||
                handRight.TrackingState == JointTrackingState.NotTracked ||
                handLeft.TrackingState == JointTrackingState.NotTracked) 
            {
                    return;
            }

            positionEllipse(ellipseHead, head, false);
            positionEllipse(ellipseHandLeft, handLeft, movementBackActive);
            positionEllipse(ellipseHandRight, handRight, movementFrontActive);

            processZoomUp(head, handRight, handLeft);

            processZoomDown(head, handRight, handLeft);

            processoMouse(head, handRight, handLeft);

            layerSeletecOculted(head, handRight, handLeft);


            if (mouse)
            {
                double i, j;
                i = 0.4 * handRight.Position.X + 0.3 * x[0] + 0.2 * x[1] + 0.1 * x[2];
                j = 0.4 * handRight.Position.Y + 0.3 * y[0] + 0.2 * y[1] + y[2] * 0.1;
                for (int t = 0; t < 2; t++)
                {
                    x[t + 1] = x[t];
                    y[t + 1] = y[t];
                }
                x[0] = handRight.Position.X + 0.3;
                y[0] = handRight.Position.Y;
                int a = (int)Math.Floor((i - originX - 0.05) * 3000);
                int b = 360 + (int)Math.Floor((j - originY + .1) * -2000);
                SetCursorPos(a, b);
                Thread.Sleep(10);
                if (handLeft.Position.Z < head.Position.Z - 0.4 && !isDragOn)
                {
                    mouse_event(MOUSEEVENTF_LEFTDOWN, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
                    isDragOn = true;
                }
                else if (!(handLeft.Position.Z < head.Position.Z - 0.4)
                          && isDragOn)
                {
                    mouse_event(MOUSEEVENTF_LEFTUP, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
                    isDragOn = false;
                }
            }
        }

        private void layerSeletecOculted(Joint head, Joint handRight, Joint handLeft)
        {
            if (handRight.Position.X > head.Position.X + 0.5 &&
                handRight.Position.Y > head.Position.Y - 0.3 &&
                handLeft.Position.X < head.Position.X - 0.4)
            {
                if (!selectLayer)
                {
                    selectLayer = true;
                    System.Windows.Forms.SendKeys.SendWait("{BACKSPACE}");
                    Thread.Sleep(milliseconds);
                }
            }
            else
            {
                selectLayer = false;
            }
        }

        private void processoMouse(Joint head, Joint handRight, Joint handLeft)
        {
            if (handLeft.Position.Y > head.Position.Y - 0.4 &&
                handLeft.Position.X > head.Position.X - 0.3 &&
                handRight.Position.X < head.Position.X + 0.5 &&
                handRight.Position.Y > head.Position.Y - 0.4)
            {
                if(!mouse)
                {
                    mouse = true;
                    originX = head.Position.X;
                    originY = head.Position.Y;
                    x[0] = x[1] = x[2] = handRight.Position.X;
                    y[0] = y[1] = y[2] = handRight.Position.Y;
                }
            }
            else
            {
                mouse = false;
            }
        }

        private void processZoomUp(Joint head, Joint handRight, Joint handLeft)
        {
            if (handRight.Position.X > head.Position.X + 0.3 && 
                handRight.Position.Y < head.Position.Y - 0.4 &&
                handLeft.Position.X < head.Position.X - 0.3 &&
                handLeft.Position.Y < head.Position.Y - 0.4)
            {
                if (!movementZoomUp)
                {
                    movementZoomUp = true;
                    System.Windows.Forms.SendKeys.SendWait("^{ADD}");
                    Thread.Sleep(milliseconds);
                }
            }
            else
            {
                movementZoomUp = false;
            }
        }

        private void processZoomDown(Joint head, Joint handRight, Joint handLeft)
        {
            if (handRight.Position.X < head.Position.X + 0.3 && 
                handRight.Position.Y < head.Position.Y - 0.4 &&
                handLeft.Position.X > head.Position.X - 0.3 &&
                handLeft.Position.Y < head.Position.Y - 0.4)
            {
                if (!movementZoomDown)
                {
                    movementZoomDown = true;
                    System.Windows.Forms.SendKeys.SendWait("^{SUBTRACT}");
                    Thread.Sleep(milliseconds);
                }
            }
            else
            {
                movementZoomDown = false;
            }
        }

        private void positionEllipse(Ellipse ellipse, Joint joint, bool activo)
        {
            if (activo)
            {
                ellipse.Width = 5;
                ellipse.Height = 5;
                ellipse.Fill = brushActivo;
            }
            else
            {
                ellipse.Width = 5;
                ellipse.Height = 5;
                ellipse.Fill = brushInactivo;
            }

            CoordinateMapper mapping = myKinect.CoordinateMapper;

            Microsoft.Kinect.SkeletonPoint vector = new Microsoft.Kinect.SkeletonPoint();

            Joint updatedJoint = new Joint();
            updatedJoint.TrackingState = JointTrackingState.Tracked;
            updatedJoint.Position = vector;

            var point = mapping.MapSkeletonPointToColorPoint(joint.Position, myKinect.ColorStream.Format);
            Canvas.SetLeft(ellipse, point.X - ellipse.Width / 2);
            Canvas.SetTop(ellipse, point.Y - ellipse.Height / 2);
        }
        void MainWindow_Unloaded(object sender, RoutedEventArgs e)
        {
            myKinect.Stop();
        }

        void myKinect_ColorFrameReady(object sender, ColorImageFrameReadyEventArgs e)
        {
            using (ColorImageFrame imagenColor = e.OpenColorImageFrame())
            {
                if (imagenColor == null)
                    return;

                if (bytesColor == null || bytesColor.Length != imagenColor.PixelDataLength)
                    bytesColor = new byte[imagenColor.PixelDataLength];

                imagenColor.CopyPixelDataTo(bytesColor);

                if (bitmapImagenColor == null)
                {
                    bitmapImagenColor = new WriteableBitmap(
                        imagenColor.Width,
                        imagenColor.Height,
                        96,
                        96,
                        PixelFormats.Bgr32,
                        null);
                }

                bitmapImagenColor.WritePixels(
                    new Int32Rect(0, 0, imagenColor.Width, imagenColor.Height),
                    bytesColor,
                    imagenColor.Width * imagenColor.BytesPerPixel,
                    0);

                imagenVideo.Source = bitmapImagenColor;
            }
        }
    }
}
