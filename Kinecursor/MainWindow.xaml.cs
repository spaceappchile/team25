using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Speech.AudioFormat;
using Microsoft.Speech.Recognition;
using Microsoft.Kinect;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;
using System.Speech.Synthesis;
using System.Xml;

namespace Kinecursor
{
    public partial class MainWindow : Window
    {
        #region Atributos
        bool isTrackingEnabled = false;
        /// <summary>
        /// Bitmap that will hold color information
        /// </summary>
        private WriteableBitmap colorBitmap;

        /// <summary>
        /// Intermediate storage for the depth data received from the camera
        /// </summary>
        private short[] depthPixels;

        /// <summary>
        /// Intermediate storage for the depth data converted to color
        /// </summary>
        private byte[] colorPixels;

        const int MaxDepthDistance = 4095; // max value returned
        const int MinDepthDistance = 800; // min value returned
        const int MaxDepthDistanceOffset = MaxDepthDistance - MinDepthDistance;

        int screenWidth = (int)SystemParameters.PrimaryScreenWidth;
        int screenHeight = (int)SystemParameters.PrimaryScreenHeight;

        bool isMouseDown = false;
        int handDepthFirst = 0;

        //hardcoded locations to Blue, Green, Red (BGR) index positions       
        const int BlueIndex = 0;
        const int GreenIndex = 1;
        const int RedIndex = 2;
        int depthX = 0;
        int depthY = 0;
        bool isMakingAFist = false;
        int lastDepthClosest = MaxDepthDistance;

        bool handReady = false;

        /// <summary>
        /// Active Kinect sensor.
        /// </summary>
        public KinectSensor sensor;

        /// <summary>
        /// Stream of audio being captured by Kinect sensor.
        /// </summary>
        public Stream audioStream;

        /// <summary>
        /// <code>true</code> if audio is currently being read from Kinect stream, <code>false</code> otherwise.
        /// </summary>
        public bool reading;

        /// <summary>
        /// Speech recognition engine using audio data from Kinect.
        /// </summary>
        public SpeechRecognitionEngine speechEngine;

        /// <summary>
        /// Thread that is reading audio from Kinect stream.
        /// </summary>
        public Thread readingThread;

        /// <summary>
        /// Synthetizer
        /// </summary>
        SpeechSynthesizer speaker = new SpeechSynthesizer();
        bool flag = false;

        /// <summary>
        /// Voz preconfigurada
        /// </summary>
        string voz = "Microsoft Zira Desktop";

        bool saludo = true;
        bool plot = true;
        bool arrastrar = false;

        #endregion

        #region Eventos de ventana
        /// <summary>
        /// Initializes a new instance of the MainWindow class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Execute startup tasks
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void WindowLoaded(object sender, RoutedEventArgs e)
        {

            foreach (var potentialSensor in KinectSensor.KinectSensors)
            {
                if (potentialSensor.Status == KinectStatus.Connected)
                {
                    this.sensor = potentialSensor;
                    break;
                }
            }

            if (null != this.sensor)
            {
                // Turn on the depth stream to receive depth frames
                this.sensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30);

                this.sensor.SkeletonStream.Enable();

                // Allocate space to put the depth pixels we'll receive
                this.depthPixels = new short[this.sensor.DepthStream.FramePixelDataLength];

                // Allocate space to put the color pixels we'll create
                this.colorPixels = new byte[this.sensor.DepthStream.FramePixelDataLength * sizeof(int)];

                // This is the bitmap we'll display on-screen
                this.colorBitmap = new WriteableBitmap(this.sensor.DepthStream.FrameWidth, this.sensor.DepthStream.FrameHeight, 96.0, 96.0, PixelFormats.Bgr32, null);

                // Set the image we display to point to the bitmap where we'll put the image data
                this.imgDepth.Source = this.colorBitmap;

                // Add an event handler to be called whenever there is new depth frame data
                this.sensor.DepthFrameReady += this.SensorDepthFrameReady;

                this.sensor.DepthStream.Range = DepthRange.Default;

                this.sensor.SkeletonFrameReady += this.SensorSkeletonFrameReady;

                this.sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Seated;

                // Start the sensor!
                try
                {
                    this.sensor.Start();


                    RecognizerInfo ri = GetKinectRecognizer();

                    if (null != ri)
                    {
                        speechEngine = new SpeechRecognitionEngine(ri.Id);

                        // Create a grammar from grammar definition XML file.
                        XmlDocument xDoc = new XmlDocument();
                        xDoc.Load(System.AppDomain.CurrentDomain.BaseDirectory + "Speech.xml");

                        //Read XML
                        using (var memoryStream = new MemoryStream(Encoding.ASCII.GetBytes(xDoc.OuterXml)))
                        {
                            var g = new Grammar(memoryStream);
                            speechEngine.LoadGrammar(g);
                        }

                        speechEngine.SpeechRecognized += SpeechRecognized;

                        // For long recognition sessions (a few hours or more), it may be beneficial to turn off adaptation of the acoustic model. 
                        // This will prevent recognition accuracy from degrading over time.
                        speechEngine.SetInputToAudioStream(
                            sensor.AudioSource.Start(), new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
                        speechEngine.RecognizeAsync(RecognizeMode.Multiple);
                    }
                }
                catch (IOException)
                {
                    this.sensor = null;
                }
            }

            if (null == this.sensor)
            {
                MessageBox.Show("No Ready Kinect Found.. :(");
            }
        }

        /// <summary>
        /// Execute shutdown tasks
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (null != this.sensor)
            {
                this.sensor.Stop();
            }
        }

        #endregion

        #region Inicializar kinect

        /// <summary>
        /// Gets the metadata for the speech recognizer (acoustic model) most suitable to
        /// process audio from Kinect device.
        /// </summary>
        /// <returns>
        /// RecognizerInfo if found, <code>null</code> otherwise.
        /// </returns>
        public RecognizerInfo GetKinectRecognizer()
        {


            foreach (RecognizerInfo recognizer in SpeechRecognitionEngine.InstalledRecognizers())
            {
                //string value;
                //recognizer.AdditionalInfo.TryGetValue("Kinect", out value);
                //if ("True".Equals(value, StringComparison.OrdinalIgnoreCase)) && "en-US".Equals(recognizer.Culture.Name, StringComparison.OrdinalIgnoreCase))
                //{
                return recognizer;
                //}
            }

            return null;
        }

        #endregion

        #region Skeleton tracking

        /// <summary>
        /// Event handler for Kinect sensor's DepthFrameReady event
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void SensorDepthFrameReady(object sender, DepthImageFrameReadyEventArgs e)
        {
            using (DepthImageFrame depthFrame = e.OpenDepthImageFrame())
            {
                if (depthFrame != null)
                {
                    // Copy the pixel data from the image to a temporary array
                    depthFrame.CopyPixelDataTo(this.depthPixels);

                    colorPixels = GenerateColoredBytes(depthFrame);

                    // Write the pixel data into our bitmap
                    this.colorBitmap.WritePixels(
                        new Int32Rect(0, 0, this.colorBitmap.PixelWidth, this.colorBitmap.PixelHeight),
                        this.colorPixels,
                        this.colorBitmap.PixelWidth * sizeof(int),
                        0);
                }
            }
        }

        /// <summary>
        /// Event handler for Kinect sensor's SkeletonFrameReady event
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void SensorSkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            Skeleton[] skeletons = new Skeleton[0];

            using (SkeletonFrame skeletonFrame = e.OpenSkeletonFrame())
            {
                if (skeletonFrame != null)
                {
                    skeletons = new Skeleton[skeletonFrame.SkeletonArrayLength];
                    skeletonFrame.CopySkeletonDataTo(skeletons);
                }
            }

            if (skeletons.Length != 0)
            {
                Skeleton skel = (from s in skeletons
                                 where s.TrackingState == SkeletonTrackingState.Tracked
                                 select s).FirstOrDefault();
                if (null == skel) { return; }

                if (skel.TrackingState == SkeletonTrackingState.Tracked && isTrackingEnabled)
                {
                    Joint jointHandLeft = skel.Joints[JointType.HandLeft];
                    Joint jointHandRight = skel.Joints[JointType.HandRight];
                    Joint jointShoulderCenter = skel.Joints[JointType.ShoulderCenter];

                    if (jointHandRight.Position.Z - jointShoulderCenter.Position.Z < -0.3)
                    {
                        float x = jointHandRight.Position.X - jointShoulderCenter.Position.X;//hX - sX;
                        float y = jointShoulderCenter.Position.Y - jointHandRight.Position.Y;//sY - hY;
                        handReady = true;
                        SetCursorPos((int)((x + 0.03) / 0.60 * screenWidth), (int)(y / 0.60 * screenHeight));
                    }
                    else if (jointHandLeft.Position.Z - jointShoulderCenter.Position.Z < -0.3)
                    {
                        float x = jointHandLeft.Position.X - jointShoulderCenter.Position.X;//hX - sX;
                        float y = jointShoulderCenter.Position.Y - jointHandLeft.Position.Y;//sY - hY;
                        handReady = true;
                        SetCursorPos((int)((x + 0.3) / 0.60 * screenWidth), (int)(y / 0.60 * screenHeight));
                    }




                    else
                    {
                        handReady = false;
                    }

                    tbInfo.Text = "Fist: " + isMakingAFist.ToString();
                }

            }

        }

        private byte[] GenerateColoredBytes(DepthImageFrame depthFrame)
        {
            //get the raw data from kinect with the depth for every pixel
            short[] rawDepthData = new short[depthFrame.PixelDataLength];
            depthFrame.CopyPixelDataTo(rawDepthData);

            //use depthFrame to create the image to display on-screen
            //depthFrame contains color information for all pixels in image
            //Height x Width x 4 (Red, Green, Blue, empty byte)
            Byte[] pixels = new byte[depthFrame.Height * depthFrame.Width * 4];


            //Bgr32  - Blue, Green, Red, empty bytel
            //Bgra32 - Blue, Green, Red, transparency 
            //You must set transparency for Bgra as .NET defaults a byte to 0 = fully transparent

            int depthClosest = MaxDepthDistance;

            List<int> lstHandTop = new List<int>();

            //loop through all distances
            //pick a RGB color based on distance
            for (int depthIndex = 0, colorIndex = 0;
                depthIndex < rawDepthData.Length && colorIndex < pixels.Length;
                depthIndex++, colorIndex += 4)
            {
                //get the player (requires skeleton tracking enabled for values)
                int player = rawDepthData[depthIndex] & DepthImageFrame.PlayerIndexBitmask;

                //gets the depth value
                int depth = rawDepthData[depthIndex] >> DepthImageFrame.PlayerIndexBitmaskWidth;

                if (player > 0)
                {

                    pixels[colorIndex + BlueIndex] = 0;
                    pixels[colorIndex + GreenIndex] = 0;
                    pixels[colorIndex + RedIndex] = 255;

                    if (depth < depthClosest)
                    {
                        depthClosest = depth;
                    }

                    //Do Finger Tracking
                    if (handReady)
                    {
                        if (depth - lastDepthClosest < 70)
                        {
                            pixels[colorIndex + BlueIndex] = 0;
                            pixels[colorIndex + GreenIndex] = 255;
                            pixels[colorIndex + RedIndex] = 255;

                            depthX = GetX(depthIndex);
                            depthY = GetY(depthIndex);

                            if (!lstHandTop.Contains(depthX))
                            {
                                if (lstHandTop.Count == 0)
                                {
                                    handDepthFirst = depth;
                                }
                                lstHandTop.Add(depthX);

                                pixels[colorIndex + BlueIndex] = 0;
                                pixels[colorIndex + GreenIndex] = 255;
                                pixels[colorIndex + RedIndex] = 0;
                            }

                        }
                    }


                }
                else
                {
                    byte intensity = CalculateIntensityFromDepth(depth);
                    pixels[colorIndex + BlueIndex] = intensity;
                    pixels[colorIndex + GreenIndex] = intensity;
                    pixels[colorIndex + RedIndex] = intensity;

                    //The last point...
                    if (IsInOnePoint(depthIndex, 319, 239))
                    {
                        lastDepthClosest = depthClosest;

                        int widthSeparator = ((MaxDepthDistance - handDepthFirst) / 65 - 20);

                        /*
                        tbInfo.Text = lstHandTop.Count.ToString() + "," + widthSeparator.ToString() + "," +
                                  (lstHandTop.Count > widthSeparator ? "True" : "False") + "," + handDepthFirst.ToString();
                    */

                        isMakingAFist = lstHandTop.Count > widthSeparator;

                        if (isMakingAFist)
                        {
                            if (!isMouseDown)
                            {
                                MouseLeftDown();
                                isMouseDown = true;

                            }
                        }
                        else
                        {
                            if (isMouseDown)
                            {
                                MouseLeftUp();
                                isMouseDown = false;
                                MouseLeftClick();

                            }
                        }
                    }
                }

            }

            return pixels;
        }

        private static byte CalculateIntensityFromDepth(int distance)
        {
            //formula for calculating monochrome intensity for histogram
            return (byte)(255 - (255 * Math.Max(distance - MinDepthDistance, 0)
                / (MaxDepthDistanceOffset)));
        }

        public int GetX(int depth)
        {
            return (depth + 1) % 320;
        }

        public int GetY(int depth)
        {
            return (depth + 1) / 320;
        }

        public bool IsInOnePoint(int depth, int x, int y)
        {
            if ((depth + 1) % 320 == x && (depth + 1) / 320 == y)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region Mouse Controll



        public void MouseLeftClick()
        {
            mouse_event(MouseEventFlag.LeftDown | MouseEventFlag.Absolute, 0, 0, 0, UIntPtr.Zero);

            mouse_event(MouseEventFlag.LeftUp | MouseEventFlag.LeftUp | MouseEventFlag.Absolute, 0, 0, 0, UIntPtr.Zero);
        }

        public void MouseLeftDown()
        {
            mouse_event(MouseEventFlag.LeftDown, 0, 0, 0, UIntPtr.Zero);
        }
        public void MouseLeftUp()
        {
            mouse_event(MouseEventFlag.LeftUp, 0, 0, 0, UIntPtr.Zero);
        }

        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);
        [DllImport("user32.dll")]
        static extern void mouse_event(MouseEventFlag flags, int dx, int dy, uint data, UIntPtr extraInfo);
        [Flags]
        enum MouseEventFlag : uint
        {
            Move = 0x0001,
            LeftDown = 0x0002,
            LeftUp = 0x0004,
            RightDown = 0x0008,
            RightUp = 0x0010,
            MiddleDown = 0x0020,
            MiddleUp = 0x0040,
            XDown = 0x0080,
            XUp = 0x0100,
            Wheel = 0x0800,
            VirtualDesk = 0x4000,
            Absolute = 0x8000
        }
        #endregion

        #region Reconocimiento
        /// <summary>
        /// Handler for recognized speech events.
        /// </summary>
        /// <param name="sender">object sending the event.</param>
        /// <param name="e">event arguments.</param>
        public void SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            try
            {

                // Speech utterance confidence below which we treat speech as if it hadn't been heard
                const double ConfidenceThreshold = 0.35;

                if (e.Result.Confidence >= ConfidenceThreshold)
                {
                    Console.WriteLine("Reconocido : " + e.Result.Confidence);
                    switch (e.Result.Semantics.Value.ToString())
                    {
                        case "plot":
                            if (plot)
                            {
                                Console.WriteLine("Plot");
                                speaker.SpeakAsync("Of course, ploting.");
                                System.Windows.Forms.SendKeys.SendWait("%");
                                System.Windows.Forms.SendKeys.SendWait("{4}");
                                System.Windows.Forms.SendKeys.SendWait("{TAB}");
                                System.Windows.Forms.SendKeys.SendWait("{TAB}");
                                System.Windows.Forms.SendKeys.SendWait("{TAB}");
                                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                                Thread.Sleep(4000);
                                speaker.SpeakAsync("Map deployed, master.");
                                Thread.Sleep(3500);
                                speaker.SpeakAsync("You're welcome.");
                                plot = false;
                            }
                            break;
                        case "hello":
                            if (saludo)
                            {
                                Console.WriteLine("Hello");
                                speaker.SpeakAsync("Hi master, Im valentina, you're personal assistent.");
                                saludo = false;
                            }
                            break;
                        case "bye":
                            if (!saludo && !plot)
                            {
                                Console.WriteLine("Bye");
                                speaker.SpeakAsync("Bye master.");
                                Thread.Sleep(3000);
                                speaker.SpeakAsync("Huston we have a problem here.");
                                Thread.Sleep(8000);
                                speaker.SpeakAsync("It's a joke  hahahah.");
                                Thread.Sleep(5000);
                                speaker.SpeakAsync("Good Bye, space apps.Thanks for that opportunity");
                                System.Windows.Forms.SendKeys.SendWait("%{F4}");
                                saludo = true;
                                plot = true;
                            }
                            break;
                        case "enable":
                            isTrackingEnabled = true;
                            Console.WriteLine("Tracking enable");
                            break;
                        case "disable":
                            isTrackingEnabled = false;
                            Console.WriteLine("Tracking disable");
                            break;
                        case "zoomIn":
                            Console.WriteLine("Zoom in");
                            System.Windows.Forms.SendKeys.SendWait("^({+})");
                            System.Windows.Forms.SendKeys.SendWait("^({+})");
                            System.Windows.Forms.SendKeys.SendWait("^({+})");
                            System.Windows.Forms.SendKeys.SendWait("^({+})");
                            speaker.SpeakAsync("Done.");
                            break;
                        case "zoomOut":
                            Console.WriteLine("Zoom out");
                            System.Windows.Forms.SendKeys.SendWait("^({-})");
                            System.Windows.Forms.SendKeys.SendWait("^({-})");
                            System.Windows.Forms.SendKeys.SendWait("^({-})");
                            System.Windows.Forms.SendKeys.SendWait("^({-})");
                            speaker.SpeakAsync("Done.");
                            break;
                        case "left":
                            Console.WriteLine("Rotate left");
                            System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                            System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                            System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                            System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                            speaker.SpeakAsync("Done.");
                            break;
                        case "right":
                            Console.WriteLine("Rotate right");
                            System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                            System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                            System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                            System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                            speaker.SpeakAsync("Done.");
                            break;
                        case "up":
                            Console.WriteLine("Rotate up");
                            System.Windows.Forms.SendKeys.SendWait("{UP}");
                            System.Windows.Forms.SendKeys.SendWait("{UP}");
                            System.Windows.Forms.SendKeys.SendWait("{UP}");
                            System.Windows.Forms.SendKeys.SendWait("{UP}");
                            speaker.SpeakAsync("Done.");
                            break;
                        case "down":
                            Console.WriteLine("Rotate down");
                            System.Windows.Forms.SendKeys.SendWait("{DOWN}");
                            System.Windows.Forms.SendKeys.SendWait("{DOWN}");
                            System.Windows.Forms.SendKeys.SendWait("{DOWN}");
                            System.Windows.Forms.SendKeys.SendWait("{DOWN}");
                            speaker.SpeakAsync("Done.");
                            break;
                    }
                }
                else
                {
                    Console.WriteLine("No reconocido : " + e.Result.Confidence);
                }

            }
            catch (Exception)
            {
                speaker.SpeakAsync("Huston, we have a problem here.");
                throw;
            }
        }

        #endregion
    }
}
