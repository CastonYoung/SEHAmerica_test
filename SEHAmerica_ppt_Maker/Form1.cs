using System;
using System.Drawing;
using System.Data;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Core;
using MO = Microsoft.Office;
using Ppt = Microsoft.Office.Interop.PowerPoint;
using Google.Apis.Customsearch.v1;
using Google.Apis.Customsearch.v1.Data;
using ListRequest = Google.Apis.Customsearch.v1.CseResource.ListRequest;

namespace SEHAmerica_ppt_Maker
{
	public partial class Form1 : Form
	{
		public const string save_err_msg = "Unexpected Error when trying to save to ";
		public const string open_err_msg = "Unexpected Error when trying to open ";
		public const string load_err_msg = "Unexpected Error when trying to load ";
		public const string powerpointformat = "Powerpoint (*.ppt;*.pptx;*.pptm)|*.ppt;*.pptx;*.pptm";
		
		Ppt.Application PowerPoint_App;
		Ppt.Presentations multi_presentations;
		Ppt.Presentation presentation;
		Ppt.CustomLayout layout;	//I went ahead and added references to all slide layouts for extensibility.
		#region ppLayouts
		Ppt.CustomLayout PpLayoutMixed			=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutMixed];			//-2 (Enum value)
		Ppt.CustomLayout PpLayoutTitle			=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTitle];			// 1
		Ppt.CustomLayout PpLayoutText			=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutText];				// 2
		Ppt.CustomLayout PpLayoutTwoColumnText	=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTwoColumnText];	// 3
		Ppt.CustomLayout PpLayoutTable			=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTable];			// 4
		Ppt.CustomLayout PpLayoutTextAndChart	=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTextAndChart];		// 5
		Ppt.CustomLayout PpLayoutChartAndText	=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutChartAndText];		// 6
		Ppt.CustomLayout PpLayoutOrgchart		=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutOrgchart];			// 7
		Ppt.CustomLayout PpLayoutChart			=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutChart];			// 8
		Ppt.CustomLayout PpLayoutTextAndClipart =>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTextAndClipart];	// 9
		Ppt.CustomLayout PpLayoutClipartAndText =>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutClipartAndText];	// 10
		Ppt.CustomLayout PpLayoutTitleOnly		=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTitleOnly];		// 11
		Ppt.CustomLayout PpLayoutBlank			=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutBlank];			// 12
		Ppt.CustomLayout PpLayoutTextAndObject	=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTextAndObject];	// 13
		Ppt.CustomLayout PpLayoutObjectAndText	=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutObjectAndText];	// 14
		Ppt.CustomLayout PpLayoutLargeObject	=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutLargeObject];		// 15
		Ppt.CustomLayout PpLayoutObject			=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutObject];			// 16
		Ppt.CustomLayout PpLayoutTextAndMediaClip=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTextAndMediaClip];	// 17
		Ppt.CustomLayout PpLayoutMediaClipAndText=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutMediaClipAndText];	// 18
		Ppt.CustomLayout PpLayoutObjectOverText =>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutObjectOverText];	// 19
		Ppt.CustomLayout PpLayoutTextOverObject =>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTextOverObject];	// 20
		Ppt.CustomLayout PpLayoutTextAndTwoObjects=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTextAndTwoObjects];// 21
		Ppt.CustomLayout PpLayoutTwoObjectsAndText=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTwoObjectsAndText];// 22
		Ppt.CustomLayout PpLayoutTwoObjectsOverText =>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTwoObjectsOverText];//23
		Ppt.CustomLayout PpLayoutFourObjects	=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutFourObjects];		// 24
		Ppt.CustomLayout PpLayoutVerticalText	=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutVerticalText];		// 25
		Ppt.CustomLayout PpLayoutClipArtAndVerticalText =>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutClipArtAndVerticalText];
		Ppt.CustomLayout PpLayoutVerticalTitleAndText =>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutVerticalTitleAndText];
		Ppt.CustomLayout PpLayoutVerticalTitleAndTextOverChart =>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutVerticalTitleAndTextOverChart];
		Ppt.CustomLayout PpLayoutTwoObjects		=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTwoObjects];		// 29
		Ppt.CustomLayout PpLayoutObjectAndTwoObjects =>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutObjectAndTwoObjects];//30
		Ppt.CustomLayout PpLayoutTwoObjectsAndObject =>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutTwoObjectsAndObject];//31
		Ppt.CustomLayout PpLayoutCustom			=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutCustom];			// 32
		Ppt.CustomLayout PpLayoutSectionHeader	=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutSectionHeader];	// 33
		Ppt.CustomLayout PpLayoutComparison		=>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutComparison];		// 34
		Ppt.CustomLayout PpLayoutContentWithCaption =>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutContentWithCaption];//35
		Ppt.CustomLayout PpLayoutPictureWithCaption =>
			presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutPictureWithCaption];//36
		#endregion
		Ppt.Slide Active_slide => PowerPoint_App.ActiveWindow.View.Slide;
		ImageList thumb_list, image_list;	//Lists of the "thumbnail" images and the full image links.
		OpenFileDialog ofd;
		string filePath = null;
		int image_count;
		bool Nopowerpoint => null == PowerPoint_App;
		MsoTriState msoTrue => MsoTriState.msoTrue;
		MsoTriState msoFalse=> MsoTriState.msoFalse;


		private struct Double_Links	//Because C# 7's Tuples suddenly decided to quit working.
		{	
			public string thumbnail;
			public string picture;
		}

		public string RemovePunct(string text, params char[] no)
		{ 
			//return RemovePunct(text.Replace(no[0], string.Empty), ((uint)no + sizeof(char)) as char[]);
			int ind = text.IndexOfAny(no);
			return ind == -1 ? text : RemovePunct(text.Remove(ind, 1), no);
		}

		public Form1()
		{
			InitializeComponent();

			image_count = 10;	//For now we're just going with 10.
			thumb_list = new ImageList();
			image_list = new ImageList();
		}

		private void CreateNewPresentation()
		{
			PowerPoint_App = new Ppt.Application();
			multi_presentations = PowerPoint_App.Presentations;
			presentation = multi_presentations.Add();
			layout = presentation.SlideMaster.CustomLayouts[Ppt.PpSlideLayout.ppLayoutText];
			presentation.Slides.AddSlide(1, layout);
		}

		private void ReadSlide(object sender, EventArgs e)
		{
			if (Nopowerpoint)						//I normally put the code on the same line as the '{' for 
			{										//if statements and loops; however, I wanted to keep the
				ofd = new OpenFileDialog();			//ofd stuff seperate. Regardless I will pay much more 
				ofd.Filter = powerpointformat;		//circumspection to style once I know the company's standards.
				try
				{	if (ofd.ShowDialog() == DialogResult.OK) filePath = ofd.FileName;
					else return;
				} catch (SystemException exc)
				{	MessageBox.Show(open_err_msg + ofd.FileName + ":\n\n" + exc);
					return;
				} finally { ofd.Dispose(); }

				PowerPoint_App = new Ppt.Application();
				multi_presentations = PowerPoint_App.Presentations;
				presentation = multi_presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
			}

			TitleBox.Text	= string.Empty;
			BodyTextBox.Text= string.Empty;

			if (layout == PpLayoutText)//Currently the program can only do PpLayoutText
			{	TitleBox.Text	= Active_slide.Shapes[1].TextFrame.TextRange.Text;
				BodyTextBox.Text= Active_slide.Shapes[2].TextFrame.TextRange.Text;
			}
			else foreach (var item in presentation.Slides[1].Shapes)
			{	var shape = (Ppt.Shape)item;
				if (shape.AlternativeText == "Click to add title")
					TitleBox.Text += shape.TextFrame.TextRange.Text + ' ';
				
				else if(shape.HasTextFrame		== MsoTriState.msoTrue
					&&	shape.TextFrame.HasText == MsoTriState.msoTrue)
					{ BodyTextBox.Text += shape.TextFrame.TextRange.Text + ' '; }
			}
		}//			^As I said ealier I'll stick to the company style guide from now on
		//			things got a little out of hand with this project.

		private void Save(object sender, EventArgs e)
		{
			if (Nopowerpoint)
			{	
				SaveFileDialog sfd = new SaveFileDialog();
				sfd.Filter = powerpointformat;
				try
				{	if (sfd.ShowDialog() == DialogResult.OK) filePath = sfd.FileName;
					else return;
				} catch (SystemException exc)
				{	MessageBox.Show(save_err_msg + sfd.FileName + ":\n\n" + exc);
					return;
				} finally { sfd.Dispose(); }

				CreateNewPresentation();
			}
			
			Active_slide.Shapes[1].TextFrame.TextRange.Text = TitleBox.Text;
			Active_slide.Shapes[2].TextFrame.TextRange.Text = BodyTextBox.Text;

			presentation.SaveAs(filePath);
		}

		private void SearchImages(object sender, EventArgs e)
		{
			if (Nopowerpoint) return;
			var imgs = GoogleImageSearch(ParseText(TitleBox.Text, BodyTextBox.Text), image_count);
			
			var pb = new PictureBox();
			pb.SizeMode = PictureBoxSizeMode.Zoom;
			for(int i = 0; i < imgs.Count; ++i)
			{	try
				{	try { pb.Load(imgs[i].thumbnail); } catch (ArgumentException) { try { pb.Load(imgs[i].picture  ); } catch (ArgumentException) { pb.Image = pb.ErrorImage; } }//Try using the picture instead of the thumbnail, or if not an error message.
					thumb_list.Images.Add(pb.Image);
					try { pb.Load(imgs[i].picture  ); } catch (ArgumentException) { try { pb.Load(imgs[i].thumbnail); } catch (ArgumentException) { pb.Image = pb.ErrorImage; } }//Try using the thumbnail instead of the picture, or if not an error message.
					image_list.Images.Add(pb.Image);
					ListView1.Items.Add(imgs[i].picture, i);
				} catch (System.Net.WebException)
				{	if (thumb_list.Images.Count == i + 1) thumb_list.Images.RemoveAt(i);
					//thumb_list.Images.Add(pb.ErrorImage);
					if (image_list.Images.Count == i + 1) image_list.Images.RemoveAt(i);
					//image_list.Images.Add(pb.ErrorImage);
				}
			}
			ListView1.SmallImageList = thumb_list;
			ListView1.LargeImageList = image_list;
			ListView1.StateImageList = image_list;
		}

		private void NewLoadImages(object sender, EventArgs e)
		{ 
			foreach (ListViewItem image in ListView1.SelectedItems)
			{	try { Active_slide.Shapes.AddPicture(image.Text, msoFalse, msoTrue, 0f, 0f); }
				catch (ArgumentException exc)
					{ MessageBox.Show(load_err_msg + image.Text + ":\n\n" + exc); }
			}
		}

		private void LoadImages()
		{	
			/*OpenFileDialog*/ ofd = new OpenFileDialog();
			ofd.Filter = "Image (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff;*.wdp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff;*.wdp";
			
			try { if (ofd.ShowDialog() != DialogResult.OK) ofd.Dispose(); }
			catch (Exception exc)
			{	MessageBox.Show(open_err_msg + ofd.FileName + ":\n\n" + exc);
				if (ofd != null) ofd.Dispose();
			}
		}

		private string[] ParseText(string titletext, string othertext)
		{	
			const StringSplitOptions nOMTs = StringSplitOptions.RemoveEmptyEntries;
			List<string> terms = new List<string>(new string[] { titletext });
			Ppt.TextRange run;

			for (int i = 1; i < image_count - 1; ++i)
			{	run = Active_slide.Shapes[2].TextFrame.TextRange.Runs(i);
				if (run == null) break;
				run.RemovePeriods();
				if (run.Font.Bold == MsoTriState.msoTrue)
				{	var punctless = RemovePunct(run.Text.Trim(), '\"','\'','?','*','\r');
					terms.AddRange(punctless.Split(new char[]{'\t','\n'}, nOMTs));
				}
			}

			return terms.ToArray();
		}

		//private List<(string /*thumbnail*/, string /*context*/)> GoogleImageSearch(string[] terms, int n_images)
		private List<Double_Links> GoogleImageSearch(string[] terms, int n_images)
		{
			var init = new Google.Apis.Services.BaseClientService.Initializer();
			init.ApiKey = "AIzaSyAX6sFmumdC70LOkqwZbGMEXr8Tcbr8Z_k";

			var searchservice = new CustomsearchService(init);

			var imageUrls = new List<Double_Links>();

			//v-I'm making the title get priority over bolded text by giving it the remainder of the n_images-v (Granted this can mean that upto half of the images are for the title.)
			Search(searchservice, terms[0], (n_images/terms.Length) + (n_images%terms.Length), /*out*/imageUrls);
			for (int i = 1; i < Math.Min(terms.Length, n_images); ++i)
				Search(searchservice, terms[i], Math.Max(1, n_images/terms.Length), /*out*/imageUrls);
			
			return imageUrls;
		}

		private void Search(CustomsearchService searchservice, string term, int n_searches, /*out*/List<Double_Links> imageUrls)
		{
			ListRequest listRequest = searchservice.Cse.List();
		//	listRequest.CreateRequest();
			listRequest.Cx = "ca0ca9bbd59600922";
			listRequest.SearchType = ListRequest.SearchTypeEnum.Image;
			listRequest.ImgColorType = ListRequest.ImgColorTypeEnum.ImgColorTypeUndefined;
			if (term == "") return;
			listRequest.Q = term;
			listRequest.Num = n_searches;
			var search = listRequest.Execute();

			if (search.Items == null) return;
				
			foreach (Result result in search.Items)
				imageUrls.Add( new Double_Links { thumbnail = result.Image.ThumbnailLink, picture = result.Link } );
		}
	}
}
