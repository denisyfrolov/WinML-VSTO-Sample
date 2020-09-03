using System;
using System.Collections.Generic;
using System.Threading;
using System.Linq;
using System.IO;
using System.Drawing;
using Windows.Storage;
using Windows.Graphics.Imaging;
using Windows.Storage.Streams;
using Windows.AI.MachineLearning;
using Windows.Foundation;
using Windows.Media;
using Newtonsoft.Json;
using System.Windows.Forms;

namespace WinML_VSTO_Sample
{
    public partial class ThisAddIn
    {
        //private static LearningModelDeviceKind _deviceKind = LearningModelDeviceKind.Default;
        private static string _deviceName = "default";
        private static string _modelPath = "Assets/SqueezeNet.onnx";
        private static string _imagePath = "Assets/kitten_224.png";
        //private static string _imagePath = "Assets/fish.png";
        private static string _labelsFileName = "Assets/Labels.json";
        private static LearningModel _model = null;
        private static LearningModelSession _session;
        private static List<string> _labels = new List<string>();
        private static RichTextBox richTextBox1;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            richTextBox1 = new RichTextBox();
            richTextBox1.Dock = DockStyle.Fill;

            richTextBox1.SelectionFont = new Font("Verdana", 12, FontStyle.Bold);
            richTextBox1.SelectionColor = Color.Red;

            Clipboard.SetImage(Image.FromFile(_imagePath));
            richTextBox1.Paste();

            // Load and create the model 
            outToLog($"Loading modelfile '{_modelPath}' on the '{_deviceName}' device");

            int ticks = Environment.TickCount;
            _model = LearningModel.LoadFromFilePath(_modelPath);
            ticks = Environment.TickCount - ticks;
            outToLog($"model file loaded in { ticks } ticks");

            // Create the evaluation session with the model and device
            _session = new LearningModelSession(_model);

            outToLog("Getting color management mode...");
            ColorManagementMode colorManagementMode = GetColorManagementMode();

            outToLog("Loading the image...");
            ImageFeatureValue imageTensor = LoadImageFile(colorManagementMode);

            // create a binding object from the session
            outToLog("Binding...");
            LearningModelBinding binding = new LearningModelBinding(_session);
            binding.Bind(_model.InputFeatures.ElementAt(0).Name, imageTensor);

            outToLog("Running the model...");
            ticks = Environment.TickCount;
            var results = _session.Evaluate(binding, "RunId");
            ticks = Environment.TickCount - ticks;
            outToLog($"model run took { ticks } ticks");

            // retrieve results from evaluation
            var resultTensor = results.Outputs[_model.OutputFeatures.ElementAt(0).Name] as TensorFloat;
            var resultVector = resultTensor.GetAsVectorView();

            PrintResults(resultVector);

            Form form1 = new Form();
            form1.Size = new Size(800, 800);
            form1.Controls.Add(richTextBox1);
            //form1.Show();
            form1.ShowDialog();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        private static void LoadLabels()
        {
            // Parse labels from label json file.  We know the file's 
            // entries are already sorted in order.
            var fileString = File.ReadAllText(_labelsFileName);
            var fileDict = JsonConvert.DeserializeObject<Dictionary<string, string>>(fileString);
            foreach (var kvp in fileDict)
            {
                _labels.Add(kvp.Value);
            }
        }


        private static T AsyncHelper<T>(IAsyncOperation<T> operation)
        {
            AutoResetEvent waitHandle = new AutoResetEvent(false);
            operation.Completed = new AsyncOperationCompletedHandler<T>((op, status) =>
            {
                waitHandle.Set();
            });
            waitHandle.WaitOne();
            return operation.GetResults();
        }

        private static ImageFeatureValue LoadImageFile(ColorManagementMode colorManagementMode)
        {
            BitmapDecoder decoder = null;
            try
            {
                FileInfo f = new FileInfo(_imagePath);
                StorageFile imageFile = AsyncHelper(StorageFile.GetFileFromPathAsync(f.FullName));
                IRandomAccessStream stream = AsyncHelper(imageFile.OpenReadAsync());
                decoder = AsyncHelper(BitmapDecoder.CreateAsync(stream));
            }
            catch (Exception e)
            {
                outToLog("Failed to load image file! Make sure that fully qualified paths are used.");
                outToLog($" Exception caught.\n {e}");
                System.Environment.Exit(e.HResult);
            }
            SoftwareBitmap softwareBitmap = null;
            try
            {
                softwareBitmap = AsyncHelper(
                    decoder.GetSoftwareBitmapAsync(
                        decoder.BitmapPixelFormat,
                        decoder.BitmapAlphaMode,
                        new BitmapTransform(),
                        ExifOrientationMode.RespectExifOrientation,
                        colorManagementMode
                    )
                );
            }
            catch (Exception e)
            {
                outToLog("Failed to create SoftwareBitmap! Please make sure that input image is within the model's colorspace.");
                outToLog($" Exception caught.\n {e}");
                System.Environment.Exit(e.HResult);
            }
            softwareBitmap = SoftwareBitmap.Convert(softwareBitmap, BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
            VideoFrame inputImage = VideoFrame.CreateWithSoftwareBitmap(softwareBitmap);
            return ImageFeatureValue.CreateFromVideoFrame(inputImage);
        }

        private static ColorManagementMode GetColorManagementMode()
        {
            // Get model color space gamma
            string gammaSpace = "";
            bool doesModelContainGammaSpaceMetadata = _model.Metadata.TryGetValue("Image.ColorSpaceGamma", out gammaSpace);
            if (!doesModelContainGammaSpaceMetadata)
            {
                outToLog("    Model does not have color space gamma information. Will color manage to sRGB by default...");
            }
            if (!doesModelContainGammaSpaceMetadata || gammaSpace.Equals("SRGB", StringComparison.CurrentCultureIgnoreCase))
            {
                return ColorManagementMode.ColorManageToSRgb;
            }
            // Due diligence should be done to make sure that the input image is within the model's colorspace. There are multiple non-sRGB color spaces.
            outToLog($"    Model metadata indicates that color gamma space is : {gammaSpace}. Will not manage color space to sRGB...");
            return ColorManagementMode.DoNotColorManage;
        }

        private static void PrintResults(IReadOnlyList<float> resultVector)
        {
            // load the labels
            LoadLabels();

            List<(int index, float probability)> indexedResults = new List<(int, float)>();
            for (int i = 0; i < resultVector.Count; i++)
            {
                indexedResults.Add((index: i, probability: resultVector.ElementAt(i)));
            }
            indexedResults.Sort((a, b) =>
            {
                if (a.probability < b.probability)
                {
                    return 1;
                }
                else if (a.probability > b.probability)
                {
                    return -1;
                }
                else
                {
                    return 0;
                }
            });

            for (int i = 0; i < 3; i++)
            {
                //Console.WriteLine($"\"{ _labels[indexedResults[i].index]}\" with confidence of { indexedResults[i].probability}");
                outToLog($"\"{ _labels[indexedResults[i].index]}\" with confidence of { indexedResults[i].probability}");
            }
        }

        static void outToLog(string output)
        {
            richTextBox1.AppendText("\r\n" + output);
            richTextBox1.ScrollToCaret();
        }
    }
}
