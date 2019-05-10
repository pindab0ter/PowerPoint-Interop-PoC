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
        readonly PowerPoint.Application application = new PowerPoint.Application();
        readonly PowerPoint.Presentation presentation;
        readonly PowerPoint.SlideShowSettings slideShowSettings;
        readonly PowerPoint.SlideShowWindow slideShowWindow;
        readonly PowerPoint.Slides slides;

        public MainWindow()
        {
            InitializeComponent();
            string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            presentation = application.Presentations.Open(path + @"\presentation.pptx", ReadOnly: MsoTriState.msoTrue, WithWindow: MsoTriState.msoFalse);
            slideShowSettings = presentation.SlideShowSettings;
            slideShowSettings.Run();

            slideShowWindow = application.SlideShowWindows[1];
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

            application.SlideShowNextSlide += Application_SlideShowNextSlide;
            cbSlides.SelectionChanged += ComboBox_SelectionChanged;
            btnBackgroundcolour.Click += BtnBackgroundcolour_Click;
        }

        private void BtnBackgroundcolour_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.Slide currentSlide = slides[slideShowWindow.View.CurrentShowPosition];
            currentSlide.FollowMasterBackground = MsoTriState.msoFalse;

            PowerPoint.ColorFormat backgroundColor = currentSlide.Background.Fill.ForeColor;

            Console.WriteLine(backgroundColor.RGB);
            if (backgroundColor.RGB == ColorTranslator.ToOle(Color.White))
            {
                backgroundColor.RGB = ColorTranslator.ToOle(Color.LightSteelBlue);
            }
            else
            {
                backgroundColor.RGB = ColorTranslator.ToOle(Color.White);
            }
        }

        private void Application_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            int currentShowPosition = Wn.View.CurrentShowPosition;
            Console.WriteLine($"User triggered next to move to the next slide ({currentShowPosition})");

            this.Dispatcher.Invoke(() =>
            {
                cbSlides.SelectedIndex = currentShowPosition - 1;
            });
        }

        private void ComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            slideShowWindow.View.GotoSlide(comboBox.SelectedIndex + 1);
        }
    }
}
