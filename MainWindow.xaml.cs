using System.Windows;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System;
using System.Windows.Controls;
using System.Drawing;
using System.IO;
using System.Reflection;

namespace PowerPoint_Interop_PoC
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        readonly PowerPoint.Application powerPoint = new PowerPoint.Application();
        readonly PowerPoint.Presentation presentation;
        readonly PowerPoint.SlideShowSettings slideShowSettings;
        readonly PowerPoint.SlideShowWindow slideShowWindow;
        readonly PowerPoint.Slides slides;

        public MainWindow()
        {
            InitializeComponent();
            string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            presentation = powerPoint.Presentations.Open(path + @"\presentation.pptx", ReadOnly: MsoTriState.msoTrue, WithWindow: MsoTriState.msoFalse);
            slideShowSettings = presentation.SlideShowSettings;
            slideShowSettings.Run();

            slideShowWindow = powerPoint.SlideShowWindows[1];
            slides = presentation.Slides;

            // Fill combobox with slides
            Dispatcher.Invoke(() =>
            {
                foreach (PowerPoint.Slide slide in slides)
                {
                    cbSlides.Items.Add(slide.Name);
                }

                cbSlides.SelectedIndex = 0;
            });

            // Subscribe to events
            cbSlides.SelectionChanged += ComboBox_SelectionChanged;
            btnBackgroundColor.Click += BtnBackgroundColor_Click;
            powerPoint.SlideShowNextSlide += OnSlideShowNextSlide;
            powerPoint.PresentationClose += OnPresentationClose;
        }

        //
        // Application Window Events
        //

        private void ComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            slideShowWindow.View.GotoSlide(comboBox.SelectedIndex + 1);
        }

        private void BtnBackgroundColor_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.Slide currentSlide = slides[slideShowWindow.View.CurrentShowPosition];
            currentSlide.FollowMasterBackground = MsoTriState.msoFalse;
            PowerPoint.ColorFormat backgroundColor = currentSlide.Background.Fill.ForeColor;

            // Toggle between colors
            if (backgroundColor.RGB == ColorTranslator.ToOle(Color.White))
            {
                backgroundColor.RGB = ColorTranslator.ToOle(Color.LightSteelBlue);
            }
            else
            {
                backgroundColor.RGB = ColorTranslator.ToOle(Color.White);
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            if (powerPoint != null)
            {
                powerPoint.Quit();
            }
            base.OnClosed(e);
        }

        //
        // PowerPoint Events
        //

        private void OnSlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            int currentShowPosition = Wn.View.CurrentShowPosition;
            Console.WriteLine($"User triggered next to move to the next slide ({currentShowPosition})");

            this.Dispatcher.Invoke(() =>
            {
                cbSlides.SelectedIndex = currentShowPosition - 1;
            });
        }

        private void OnPresentationClose(PowerPoint.Presentation Pres)
        {
            Pres.Saved = MsoTriState.msoTrue;
            Pres.RejectAll();
            this.Dispatcher.Invoke(() =>
            {
                Application.Current.Shutdown();
            });
        }
    }
}
