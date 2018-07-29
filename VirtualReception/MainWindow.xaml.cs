// 
// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license.
// 
// Microsoft Cognitive Services: http://www.microsoft.com/cognitive
// 
// Microsoft Cognitive Services Github:
// https://github.com/Microsoft/Cognitive
// 
// Copyright (c) Microsoft Corporation
// All rights reserved.
// 
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
// 
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// 

using System;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Newtonsoft.Json;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Common = Microsoft.ProjectOxford.Common;
using FaceAPI = Microsoft.ProjectOxford.Face;
using VisionAPI = Microsoft.ProjectOxford.Vision;
using System.IO;
using System.Drawing.Imaging;
using Microsoft.CognitiveServices.SpeechRecognition;
using System.Configuration;
using System.Speech.Synthesis;
using SpeechToTextWPFSample;
using System.Threading;
using VideoFrameAnalyzer;

namespace LiveCameraSample
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        #region Video
        private FaceAPI.FaceServiceClient _faceClient = null;
        private VisionAPI.VisionServiceClient _visionClient = null;
        private readonly FrameGrabber<LiveCameraResult> _grabber = null;
        private static readonly ImageEncodingParam[] s_jpegParams = {
            new ImageEncodingParam(ImwriteFlags.JpegQuality, 60)
        };
        private readonly CascadeClassifier _localFaceDetector = new CascadeClassifier();
        private bool _fuseClientRemoteResults;
        private LiveCameraResult _latestResultsToDisplay = null;      
        private DateTime _startTime;
        private bool ImageSaveInProgress = false;
        private bool MessageInProgress = false;


        public MainWindow()
        {
            InitializeComponent();

            // Create grabber. 
            _grabber = new FrameGrabber<LiveCameraResult>();

            // Set up a listener for when the client receives a new frame.
            _grabber.NewFrameProvided += (s, e) =>
            {
                // The callback may occur on a different thread, so we must use the
                // MainWindow.Dispatcher when manipulating the UI. 
                this.Dispatcher.BeginInvoke((Action)(() =>
                {
                    // Display the image in the left pane.
                    LeftImage.Source = e.Frame.Image.ToBitmapSource();
                }));

            };

            // Set up a listener for when the client receives a new result from an API call. 
            _grabber.NewResultAvailable += (s, e) =>
            {

                this.Dispatcher.BeginInvoke((Action)(() =>
                {
                    if (e.TimedOut)
                    {
                        MessageArea.Text = "API call timed out.";
                    }
                    else if (e.Exception != null)
                    {
                        string apiName = "";
                        string message = e.Exception.Message;
                        var faceEx = e.Exception as FaceAPI.FaceAPIException;
                        var emotionEx = e.Exception as Common.ClientException;
                        var visionEx = e.Exception as VisionAPI.ClientException;
                        if (faceEx != null)
                        {
                            apiName = "Face";
                            message = faceEx.ErrorMessage;
                        }
                        MessageArea.Text = string.Format("{0} API call failed on frame {1}. Exception: {2}", apiName, e.Frame.Metadata.Index, message);
                    }
                    else
                    {
                        _latestResultsToDisplay = e.Analysis;

                        MessageArea.Text = "";
                        if (_latestResultsToDisplay.Faces.Count() == 0)
                        {

                            LookIntoCamera();
                            //MessageArea.Text = "Look into the camera please ";
                        }
                        else
                        {
                            _latestResultsToDisplay = e.Analysis;

                            MessageArea.Text = "";
                            if (_latestResultsToDisplay.Faces.Count() == 0)
                            {
                                //MessageArea.Text = "Look into the camera please ";
                            }
                            else
                            {
                                MessageArea.Text = "";

                                if (_latestResultsToDisplay.EmotionScores != null && _latestResultsToDisplay.EmotionScores.Count() > 0)
                                {
                                    if (_latestResultsToDisplay.EmotionScores[0].Happiness <= 0.5)
                                    {
                                        //if(ImageSaveInProgress!=true)
                                        MessageArea.Text = "Smile Please ! ";
                                        //TextToSpeechNow(MessageClass.Smile,true);

                                    }
                                    else
                                    {
                                        //MessageArea.Text = "";
                                        if (ImageSaveInProgress == false)
                                        {
                                            ImageSaveInProgress = true;
                                            BitmapSource visImage = e.Frame.Image.ToBitmapSource();
                                            SaveImage(visImage, e.Frame);
                                        }
                                    }
                                }
                            }
                        }

                    }
                }));
            };

        }


        public void LookIntoCamera()
        {
            TextToSpeechNow(MessageConstants.PleaseStandInfront, true);
        }

        private void SaveImage(BitmapSource bitmapsource, VideoFrame vf)
        {
            //Bitmap imageToSave = BitmapFromSource(bitmapsource);
            MemoryStream imgData = vf.Image.ToMemoryStream(".jpg", s_jpegParams);
            System.Drawing.Image img = System.Drawing.Image.FromStream(imgData);
            img.Save(@"C:\\temp\\" + "myImage.Jpeg", ImageFormat.Jpeg);
            //MessageArea.Text = "Thanks Image Grabbed ";           
            TextToSpeechNow(MessageConstants.ThankYOU, true);
            CameraThreadStop();
        }

        private async void CameraThreadStop()
        {
            await this.Dispatcher.Invoke(async () =>
            {
                LeftImage.Visibility = Visibility.Hidden;
                await _grabber.StopProcessingAsync();
                // _grabber.StopCameraThread();

            });

        }

        /// <summary> Function which submits a frame to the Emotion API. </summary>
        /// <param name="frame"> The video frame to submit. </param>
        /// <returns> A <see cref="Task{LiveCameraResult}"/> representing the asynchronous API call,
        ///     and containing the emotions returned by the API. </returns>
        private async Task<LiveCameraResult> EmotionAnalysisFunction(VideoFrame frame)
        {
            // Encode image. 
            var jpg = frame.Image.ToMemoryStream(".jpg", s_jpegParams);
            // Submit image to API. 
            FaceAPI.Contract.Face[] faces = null;

            // See if we have local face detections for this image.
            var localFaces = (OpenCvSharp.Rect[])frame.UserData;
            if (localFaces == null || localFaces.Count() > 0)
            {
                // If localFaces is null, we're not performing local face detection.
                // Use Cognigitve Services to do the face detection.
                Properties.Settings.Default.FaceAPICallCount++;
                faces = await _faceClient.DetectAsync(
                    jpg,
                    /* returnFaceId= */ false,
                    /* returnFaceLandmarks= */ false,
                    new FaceAPI.FaceAttributeType[1] { FaceAPI.FaceAttributeType.Emotion });
            }
            else
            {
                // Local face detection found no faces; don't call Cognitive Services.
                faces = new FaceAPI.Contract.Face[0];
            }

            // Output. 
            return new LiveCameraResult
            {
                Faces = faces.Select(e => CreateFace(e.FaceRectangle)).ToArray(),
                // Extract emotion scores from results. 
                EmotionScores = faces.Select(e => e.FaceAttributes.Emotion).ToArray()
            };
        }

       

       
        private BitmapSource VisualizeResult(VideoFrame frame)
        {
            // Draw any results on top of the image. 
            BitmapSource visImage = frame.Image.ToBitmapSource();

            var result = _latestResultsToDisplay;

            if (result != null)
            {
                // See if we have local face detections for this image.
                var clientFaces = (OpenCvSharp.Rect[])frame.UserData;
                if (clientFaces != null && result.Faces != null)
                {
                    // If so, then the analysis results might be from an older frame. We need to match
                    // the client-side face detections (computed on this frame) with the analysis
                    // results (computed on the older frame) that we want to display. 
                    MatchAndReplaceFaceRectangles(result.Faces, clientFaces);
                }

                //visImage = Visualization.DrawFaces(visImage, result.Faces, result.EmotionScores, result.CelebrityNames);
               // visImage = Visualization.DrawTags(visImage, result.Tags);
            }

            return visImage;
        }

        /// <summary> Populate CameraList in the UI, once it is loaded. </summary>
        /// <param name="sender"> Source of the event. </param>
        /// <param name="e">      Routed event information. </param>
        private void CameraList_Loaded(object sender, RoutedEventArgs e)
        {
            int numCameras = _grabber.GetNumCameras();

            if (numCameras == 0)
            {
                MessageArea.Text = "No cameras found!";
            }

            var comboBox = sender as ComboBox;
            comboBox.ItemsSource = Enumerable.Range(0, numCameras).Select(i => string.Format("Camera {0}", i + 1));
            comboBox.SelectedIndex = 0;
        }

        /// <summary> Populate ModeList in the UI, once it is loaded. </summary>
        /// <param name="sender"> Source of the event. </param>
        /// <param name="e">      Routed event information. </param>
        private void ModeList_Loaded(object sender, RoutedEventArgs e)
        {
            // var modes = (AppMode[])Enum.GetValues(typeof(AppMode));

            var comboBox = sender as ComboBox;
            comboBox.ItemsSource = "Emotions";// modes.Select(m => m.ToString());
            comboBox.SelectedIndex = 0;
        }

        private void ModeList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _grabber.AnalysisFunction = EmotionAnalysisFunction;
        }

        private async void CameraThreadStart()
        {

            this.Dispatcher.Invoke(async () =>
            {
                if (!CameraList.HasItems)
                {
                    MessageArea.Text = "No cameras found; cannot start processing";
                    return;
                }

                // Clean leading/trailing spaces in API keys. 
                Properties.Settings.Default.FaceAPIKey = ConfigurationManager.AppSettings["FaceAPIKey"].ToString();
                //Properties.Settings.Default.VisionAPIKey = Properties.Settings.Default.VisionAPIKey.Trim();

                // Create API clients. 
                _faceClient = new FaceAPI.FaceServiceClient(Properties.Settings.Default.FaceAPIKey, ConfigurationManager.AppSettings["FaceAPIEndPoint"].ToString());
                // _visionClient = new VisionAPI.VisionServiceClient(Properties.Settings.Default.VisionAPIKey, Properties.Settings.Default.VisionAPIHost);

                // How often to analyze. 
                _grabber.TriggerAnalysisOnInterval(Properties.Settings.Default.AnalysisInterval);

                // Reset message. 
                MessageArea.Text = "";

                // Record start time, for auto-stop
                _startTime = DateTime.Now;

                await _grabber.StartProcessingCameraAsync(CameraList.SelectedIndex);
            });
           
        }

        private async void StartButton_Click(object sender, RoutedEventArgs e)
        {
            MicroPhoneThread();
        }

        private async void StopButton_Click(object sender, RoutedEventArgs e)
        {
            await _grabber.StopProcessingAsync();
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            //SettingsPanel.Visibility = 1 - SettingsPanel.Visibility;
        }

        private void SaveSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            //SettingsPanel.Visibility = Visibility.Hidden;
            Properties.Settings.Default.Save();
        }

        

        private FaceAPI.Contract.Face CreateFace(FaceAPI.Contract.FaceRectangle rect)
        {
            return new FaceAPI.Contract.Face
            {
                FaceRectangle = new FaceAPI.Contract.FaceRectangle
                {
                    Left = rect.Left,
                    Top = rect.Top,
                    Width = rect.Width,
                    Height = rect.Height
                }
            };
        }

        private FaceAPI.Contract.Face CreateFace(VisionAPI.Contract.FaceRectangle rect)
        {
            return new FaceAPI.Contract.Face
            {
                FaceRectangle = new FaceAPI.Contract.FaceRectangle
                {
                    Left = rect.Left,
                    Top = rect.Top,
                    Width = rect.Width,
                    Height = rect.Height
                }
            };
        }

        private FaceAPI.Contract.Face CreateFace(Common.Rectangle rect)
        {
            return new FaceAPI.Contract.Face
            {
                FaceRectangle = new FaceAPI.Contract.FaceRectangle
                {
                    Left = rect.Left,
                    Top = rect.Top,
                    Width = rect.Width,
                    Height = rect.Height
                }
            };
        }

        private void MatchAndReplaceFaceRectangles(FaceAPI.Contract.Face[] faces, OpenCvSharp.Rect[] clientRects)
        {
            // Use a simple heuristic for matching the client-side faces to the faces in the
            // results. Just sort both lists left-to-right, and assume a 1:1 correspondence. 

            // Sort the faces left-to-right. 
            var sortedResultFaces = faces
                .OrderBy(f => f.FaceRectangle.Left + 0.5 * f.FaceRectangle.Width)
                .ToArray();

            // Sort the clientRects left-to-right.
            var sortedClientRects = clientRects
                .OrderBy(r => r.Left + 0.5 * r.Width)
                .ToArray();

            // Assume that the sorted lists now corrrespond directly. We can simply update the
            // FaceRectangles in sortedResultFaces, because they refer to the same underlying
            // objects as the input "faces" array. 
            for (int i = 0; i < Math.Min(faces.Length, clientRects.Length); i++)
            {
                // convert from OpenCvSharp rectangles
                OpenCvSharp.Rect r = sortedClientRects[i];
                sortedResultFaces[i].FaceRectangle = new FaceAPI.Contract.FaceRectangle { Left = r.Left, Top = r.Top, Width = r.Width, Height = r.Height };
            }
        }
        

        #endregion







        #region Speech

        /// <summary>
        /// The isolated storage subscription key file name.
        /// </summary>
        private const string IsolatedStorageSubscriptionKeyFileName = "Subscription.txt";

        /// <summary>
        /// The default subscription key prompt message
        /// </summary>
        private const string DefaultSubscriptionKeyPromptMessage = "Paste your subscription key here to start";

        /// <summary>
        /// You can also put the primary key in app.config, instead of using UI.
        /// string subscriptionKey = ConfigurationManager.AppSettings["primaryKey"];
        /// </summary>
        private string subscriptionKey;

        /// <summary>
        /// The data recognition client
        /// </summary>
        private DataRecognitionClient dataClient;

        /// <summary>
        /// The microphone client
        /// </summary>
        private MicrophoneRecognitionClient micClient;

           
        /// <summary>
        /// Gets or sets subscription key
        /// </summary>
        public string SubscriptionKey
        {
            get
            {
                return ConfigurationManager.AppSettings["subsKeySpeech"].ToString(); ;
            }          
        }

        /// <summary>
        /// Gets the LUIS endpoint URL.
        /// </summary>
        /// <value>
        /// The LUIS endpoint URL.
        /// </value>
        private string LuisEndpointUrl
        {
            get { return ConfigurationManager.AppSettings["LuisEndpointUrl"]; }
        }

        /// <summary>
        /// Gets the default locale.
        /// </summary>
        /// <value>
        /// The default locale.
        /// </value>
        private string DefaultLocale
        {
            get { return "en-US"; }
        }

        /// <summary>
        /// Gets the Cognitive Service Authentication Uri.
        /// </summary>
        /// <value>
        /// The Cognitive Service Authentication Uri.  Empty if the global default is to be used.
        /// </value>
        private string AuthenticationUri
        {
            get
            {
                return ConfigurationManager.AppSettings["AuthenticationUri"];
            }
        }
        
        /// <summary>
        /// Handles the Click event of the _startButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        public void MicroPhoneThread()
        {
            //this.LogRecognitionStart();

            if (true)
            {
                if (this.micClient == null)
                {
                    if (true)
                    {
                        this.CreateMicrophoneRecoClientWithIntent();
                    }                    
                }

                this.micClient.StartMicAndRecognition();
            }           
        }
        
        /// <summary>
        /// Creates a new microphone reco client with LUIS intent support.
        /// </summary>
        private void CreateMicrophoneRecoClientWithIntent()
        {
            this.WriteLine("--- Start microphone dictation with Intent detection ----");

            this.micClient =
                SpeechRecognitionServiceFactory.CreateMicrophoneClientWithIntentUsingEndpointUrl(
                    this.DefaultLocale,
                    this.SubscriptionKey,
                    this.LuisEndpointUrl);
            this.micClient.AuthenticationUri = this.AuthenticationUri;
            this.micClient.OnIntent += this.OnIntentHandler;

            // Event handlers for speech recognition results
            this.micClient.OnMicrophoneStatus += this.OnMicrophoneStatus;
            this.micClient.OnPartialResponseReceived += this.OnPartialResponseReceivedHandler;
            this.micClient.OnResponseReceived += this.OnMicShortPhraseResponseReceivedHandler;
            this.micClient.OnConversationError += this.OnConversationErrorHandler;
        }


        /// <summary>
        /// Called when a final response is received;
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="SpeechResponseEventArgs"/> instance containing the event data.</param>
        private void OnMicShortPhraseResponseReceivedHandler(object sender, SpeechResponseEventArgs e)
        {
            try
            {
                Dispatcher.Invoke((Action)(() =>
                {
                    this.WriteLine("--- OnMicShortPhraseResponseReceivedHandler ---");

                    // we got the final result, so it we can end the mic reco.  No need to do this
                    // for dataReco, since we already called endAudio() on it as soon as we were done
                    // sending all the data.
                   // this.micClient.EndMicAndRecognition();

                    this.WriteResponseResult(e);

                    // MicroPhoneThread();
                    //_startButton.IsEnabled = true;
                    // _radioGroup.IsEnabled = true;
                }));
            }
            catch(Exception ex)
            {

            }
            
        }
           
        /// <summary>
        /// Writes the response result.
        /// </summary>
        /// <param name="e">The <see cref="SpeechResponseEventArgs"/> instance containing the event data.</param>
        private void WriteResponseResult(SpeechResponseEventArgs e)
        {
            if (e.PhraseResponse.Results.Length == 0)
            {
                this.WriteLine("No phrase response is available.");
                MicroPhoneThread();
            }
            else
            {
                this.WriteLine("********* Final n-BEST Results *********");
                for (int i = 0; i < e.PhraseResponse.Results.Length; i++)
                {
                    this.WriteLine(
                        "[{0}] Confidence={1}, Text=\"{2}\"",
                        i,
                        e.PhraseResponse.Results[i].Confidence,
                        e.PhraseResponse.Results[i].DisplayText);
                }

                this.WriteLine();
            }
        }
  
        /// <summary>
        /// Called when a final response is received and its intent is parsed
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="SpeechIntentEventArgs"/> instance containing the event data.</param>
        private void OnIntentHandler(object sender, SpeechIntentEventArgs e)
        {
            //this.WriteLine("--- Intent received by OnIntentHandler() ---");
            //this.WriteLine("{0}", e.Payload);            
            var _Data = JsonConvert.DeserializeObject<LUISResponse>(e.Payload);
            HandleIntent(_Data);
            //var employeeName = json["entities"].Count() != 0 ? (string)json["entities"][0].First.First : "";
            //var MeetingType = json["entities"].Count() != 0 ? (string)json["entities"][1].First.First : "";
            //var entityFound = _Data.entities.Count()>0 ? _Data.entities[0].entity : "";
            // var topIntent = _Data.intents.Count()>0 ?  _Data.intents[0].intent : "";
            /* if (entityFound != "" && topIntent!="")
             {
                 TextToSpeechNow("right , so do you have "+ topIntent + " with " + entityFound);
             }
             else
             {
                 TextToSpeechNow("sorry i dont understand you. Who do you need to see here?");
             }*/

            this.WriteLine();
        }


        private void HandleIntent(LUISResponse _Data)
        {
            this.micClient.EndMicAndRecognition();

            var entityFound = _Data.entities.Count()>0 ? _Data.entities[0].entity : "";
            var topIntent = _Data.intents.Count()>0 ?  _Data.intents[0].intent : "";

            switch (topIntent)
            {
                case "Meeting":
                    MeetingContext();
                    //FindNextMessage();
                   // MicroPhoneThread();
                    break;
                case "VisitorInfo":
                    VisitorInfo(entityFound);
                   //FindNextMessage();
                 //   MicroPhoneThread();
                    break;
                case "Consent":
                    Consent(_Data);
                    //FindNextMessage();
                    //MicroPhoneThread();
                    break;
                case "None":
                    TextToSpeechNow("sorry i dont understand you",false);
                    //MicroPhoneThread();
                   // FindNextMessage();
                    
                    break;
                default:
                    //Console.WriteLine("Default case");
                    break;
            }

            FindNextMessage();
        }

        public void FindNextMessage()
        {
            if (!LiveCameraSample.VisitorInfo.UserContextTaken)
            {
                TextToSpeechNow(LiveCameraSample.MessageConstants.MEETINGMESSAGE,false);
                MicroPhoneThread();
                return;
            }

            if (string.IsNullOrWhiteSpace(LiveCameraSample.VisitorInfo.LastName))
            {
                TextToSpeechNow(LiveCameraSample.MessageConstants.LASTNAME, false);
                MicroPhoneThread();
                return;
            }

            if (string.IsNullOrWhiteSpace(LiveCameraSample.VisitorInfo.firstName))
            {
                TextToSpeechNow(LiveCameraSample.MessageConstants.FIRSTNAME, false);
                MicroPhoneThread();
                return;
            }


            if (!LiveCameraSample.VisitorInfo.PictureTaken)
            {
                TextToSpeechNow(LiveCameraSample.MessageConstants.PictrueConsent, false);
                MicroPhoneThread();
                return;
            }

            //this.micClient.EndMicAndRecognition();

        }

        public void VisitorInfo(string entity)
        {
            if (string.IsNullOrWhiteSpace(LiveCameraSample.VisitorInfo.LastName))
            {
                LiveCameraSample.VisitorInfo.LastName = entity;
                return;
            }

            if (string.IsNullOrWhiteSpace(LiveCameraSample.VisitorInfo.firstName))
            {
                LiveCameraSample.VisitorInfo.firstName = entity;
                return;
            }

           
        }

        private void Consent(LUISResponse _Data)
        {
            var entityFound = _Data.entities.Count() > 0 ? _Data.entities[0].entity : "";
            var topIntent = _Data.intents.Count() > 0 ? _Data.intents[0].intent : "";
            var sentiment = _Data.sentimentAnalysis !=null && _Data.sentimentAnalysis.label.ToLower() =="positive" &&  _Data.sentimentAnalysis.score > Convert.ToDouble(0.80) ? "Positive" : "";

            if (!string.IsNullOrEmpty(LiveCameraSample.VisitorInfo.firstName) && !string.IsNullOrEmpty(LiveCameraSample.VisitorInfo.LastName) &&  !LiveCameraSample.VisitorInfo.PictureTaken && sentiment =="Positive")
            {
                this.micClient.EndMicAndRecognition();
                CameraThreadStart();
                LiveCameraSample.VisitorInfo.PictureTaken = true;
               
                return;
            }
            else
            {
                LiveCameraSample.VisitorInfo.PictureTaken = false;
                //FindNextMessage();
                return;
            }
        }

        public void MeetingContext()
        {
            LiveCameraSample.VisitorInfo.UserContextTaken = true;
        }

        public void TextToSpeechNow(string text,bool NewThread)
        {
            if (MessageInProgress == false)
            {
                MessageInProgress = true;
                if(NewThread)
                {
                    new Thread(() =>
                    {
                        SpeechSynthesizer synthesizer = new SpeechSynthesizer();
                        synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Teen);
                        synthesizer.Volume = 100;  // 0...100
                        synthesizer.Rate = -1;     // -10...10

                        synthesizer.Speak(text);
                        MessageInProgress = false;
                        // }

                    }).Start();
                }
                else
                {
                    SpeechSynthesizer synthesizer = new SpeechSynthesizer();
                    synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Teen);
                    synthesizer.Volume = 100;  // 0...100
                    synthesizer.Rate = -1;     // -10...10

                    synthesizer.Speak(text);
                    MessageInProgress = false;

                }
                
            }
        }
        /// <summary>
        /// Called when a partial response is received.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="PartialSpeechResponseEventArgs"/> instance containing the event data.</param>
        private void OnPartialResponseReceivedHandler(object sender, PartialSpeechResponseEventArgs e)
        {
            //this.WriteLine("--- Partial result received by OnPartialResponseReceivedHandler() ---");
            //this.WriteLine("{0}", e.PartialResult);
            //this.WriteLine();
        }

        /// <summary>
        /// Called when an error is received.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="SpeechErrorEventArgs"/> instance containing the event data.</param>
        private void OnConversationErrorHandler(object sender, SpeechErrorEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                //_startButton.IsEnabled = true;
                //_radioGroup.IsEnabled = true;
            });

            this.WriteLine("--- Error received by OnConversationErrorHandler() ---");
            this.WriteLine("Error code: {0}", e.SpeechErrorCode.ToString());
            this.WriteLine("Error text: {0}", e.SpeechErrorText);
            this.WriteLine();
        }

        /// <summary>
        /// Called when the microphone status has changed.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="MicrophoneEventArgs"/> instance containing the event data.</param>
        private void OnMicrophoneStatus(object sender, MicrophoneEventArgs e)
        {
            try {
                Dispatcher.Invoke(() =>
                {
                    WriteLine("--- Microphone status change received by OnMicrophoneStatus() ---");
                    WriteLine("********* Microphone status: {0} *********", e.Recording);
                    if (e.Recording)
                    {
                        WriteLine("Please start speaking.");
                    }

                    WriteLine();
                });
            }
            catch(Exception ex)
            {

            }
            
        }

        /// <summary>
        /// Writes the line.
        /// </summary>
        private void WriteLine()
        {
            this.WriteLine(string.Empty);
        }

        /// <summary>
        /// Writes the line.
        /// </summary>
        /// <param name="format">The format.</param>
        /// <param name="args">The arguments.</param>
        private void WriteLine(string format, params object[] args)
        {
            try{
                var formattedStr = string.Format(format, args);
                Trace.WriteLine(formattedStr);
                Dispatcher.Invoke(() =>
                {
                    _logText.Text += (formattedStr + "\n");
                    _logText.ScrollToEnd();
                });
            } catch(Exception ex)
            {

            }
           
        }
             
        #endregion
    }
}
