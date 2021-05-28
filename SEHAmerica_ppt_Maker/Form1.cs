﻿using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Core;
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
		public const string frOz_err_msg = "PowerPoint is in an unresponsive state.  Please resolve.";
		public const string powerpointformat = "Powerpoint (*.ppt;*.pptx;*.pptm)|*.ppt;*.pptx;*.pptm";
		
		Ppt.Application PowerPoint_App;
		Ppt.Presentations multi_presentations;
		Ppt.Presentation presentation;
		Ppt.PpSlideLayout layout;	//I went ahead and added references to all slide layouts for extensibility.
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
		ImageList thumb_list, image_list;
		OpenFileDialog ofd;
		List<string> box_list;
		string file_path = null;
		int image_count;

		MsoTriState msoTrue => MsoTriState.msoTrue;
		MsoTriState msoFalse=> MsoTriState.msoFalse;


		private struct Double_Links	//Because C# 7's Tuples suddenly decided to quit working.
		{	
			public string thumbnail;
			public string picture;
		}

		public string RemovePunct(string text, params char[] punct)
		{ 
			//return RemovePunct(text.Replace(no[0], string.Empty), ((uint)no + sizeof(char)) as char[]);
			int ind = text.IndexOfAny(punct);
			return ind == -1 ? text : RemovePunct(text.Remove(ind, 1), punct);
		}

		public Form1()
		{
			InitializeComponent();

			image_count = 10;	//For now we're just going with 10.
			thumb_list = new ImageList();
			image_list = new ImageList
				{ ImageSize = new System.Drawing.Size(160, 160) };
			box_list = new List<string>();
		}

		public bool NoPowerPoint()
		{	
			try { return null == PowerPoint_App || "PowerPoint" == PowerPoint_App.Caption; }
			catch(System.Runtime.InteropServices.COMException)
			{	MessageBox.Show(frOz_err_msg);
				return true;
			}
		}

		private void CreateNewPresentation()
		{
			PowerPoint_App = new Ppt.Application();
			multi_presentations = PowerPoint_App.Presentations;
			presentation = multi_presentations.Add();
			layout = Ppt.PpSlideLayout.ppLayoutText;
			presentation.Slides.AddSlide(1, presentation.SlideMaster.CustomLayouts[layout]);
		}
		
		private void ReadSlide(object sender, EventArgs e)
		{
			if (NoPowerPoint())
			{	ofd = new OpenFileDialog { Filter = powerpointformat };
				try
				{	if (ofd.ShowDialog() == DialogResult.OK) file_path = ofd.FileName;
					else return;
				} catch (SystemException exc)
				{	MessageBox.Show(open_err_msg + ofd.FileName + ":\n\n" + exc);
					return;
				} finally { ofd.Dispose(); }

				PowerPoint_App = new Ppt.Application();
				multi_presentations = PowerPoint_App.Presentations;
				presentation = multi_presentations.Open(file_path, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
			}

			TitleBox.Text	= string.Empty;
			BodyTextBox.Text= string.Empty;
			layout = Active_slide.Layout;

			if (layout == Ppt.PpSlideLayout.ppLayoutText)
			{	TitleBox.Text	= Active_slide.Shapes[1].TextFrame.TextRange.Text;
				BodyTextBox.Text= Active_slide.Shapes[2].TextFrame.TextRange.Text;
			}
			else foreach (var item in presentation.Slides[1].Shapes)
			{	var shape = (Ppt.Shape)item;
				if (shape.HasTextFrame != MsoTriState.msoTrue) continue;

				if (shape.Name.Contains("Title"))
				{	TitleBox.Text += shape.TextFrame.TextRange.Text + ' ';
					box_list.Insert(0, shape.Name);
				}
				
				else if(shape.TextFrame.HasText == MsoTriState.msoTrue)
				{	IDataObject clipped = Clipboard.GetDataObject();
					shape.TextFrame.TextRange.Copy();
					BodyTextBox.Paste();
					Clipboard.SetDataObject(clipped, true);
					box_list.Add(shape.Name);
				}
			}
		}

		private void Save(object sender, EventArgs e)
		{
			try
			{	if (NoPowerPoint())
				{	SaveFileDialog sfd = new SaveFileDialog { Filter = powerpointformat };
					try
					{	if (sfd.ShowDialog() == DialogResult.OK) file_path = sfd.FileName;
						else return;
					} catch (SystemException exc)
					{	MessageBox.Show(save_err_msg + sfd.FileName + ":\n\n" + exc);
						return;
					} finally { sfd.Dispose(); }
				
					CreateNewPresentation();
				}
				else foreach (var item in presentation.Slides[1].Shapes)
				{	var shape = (Ppt.Shape)item;
					if (shape.HasTextFrame != MsoTriState.msoTrue) continue;

					if (shape.Name == box_list[0])
						shape.TextFrame.TextRange.Text = TitleBox.Text;
				
					else if(box_list.Contains(shape.Name))
					{	IDataObject clipped = Clipboard.GetDataObject();
						BodyTextBox.SelectAll();
						BodyTextBox.Copy();
						shape.TextFrame.TextRange.Paste();
						Clipboard.SetDataObject(clipped, true);
					}
				}

				presentation.SaveAs(file_path);

			} catch (ArgumentException exc)
			{	MessageBox.Show(save_err_msg + file_path + ":\n\n" + exc);	}
		}

		private void SearchImages(object sender, EventArgs e)
		{
			if (NoPowerPoint()) return;

			var imgs = GoogleImageSearch(ParseText(TitleBox.Text, BodyTextBox.Rtf), image_count);
			
			var pb = new PictureBox { SizeMode = PictureBoxSizeMode.Zoom };
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

		private string[] ParseText(string titletext, string richtext)
		{	
			const StringComparison matchKs = StringComparison.Ordinal;
			List<string> innerterms = new List<string>(new string[] { titletext });
			List<int> startindices = new List<int>(), endindices = new List<int>{-3};

			
			try
			{	for (int i = 0; i < image_count - 1; ++i)
				{	var starti = richtext.IndexOf("\\b" , endindices[i] + 3, matchKs);
					char after_b = richtext[starti + 2];
					while (char.IsLetterOrDigit(after_b) && after_b != '1')// Old versions of the Rtf standard used \b1 for turning bold on.
					{	starti = richtext.IndexOf("\\b" , starti + 2, matchKs);
						if (-1 == starti) break;
						after_b = richtext[starti + 2];
					}
					if (-1 == starti) break;
					startindices.Add(richtext.IndexOf(" ", starti, matchKs));
					var endi = richtext.IndexOf("\\b0" , startindices[i] + 1, matchKs);
					if (-1 == endi) break;
					endindices.Add(endi);
				}
			} catch(IndexOutOfRangeException) { }
			
			for (int i = 0; i < startindices.Count; ++i)
			{	int begin = startindices[i];
				var boldedtext = richtext.Substring(begin, endindices[i+1] - begin);
				var filteredtext = RemovePunct(boldedtext, '.', '\"','\'','?','*','\r').Trim();
				if (string.Empty != filteredtext) innerterms.Add(filteredtext);
			}

			return innerterms.ToArray();
		}

		private List<Double_Links> GoogleImageSearch(string[] terms, int n_images)
		{
			var init = new Google.Apis.Services.BaseClientService.Initializer
				{ ApiKey = "AIzaSyAX6sFmumdC70LOkqwZbGMEXr8Tcbr8Z_k" };

			var searchservice = new CustomsearchService(init);

			var imageUrls = new List<Double_Links>();

			//v-I'm making the title get priority over bolded text by giving it the remainder of the n_images-v (Granted this can mean that upto half of the images are for the title.)
			Search(searchservice, terms[0], (n_images/terms.Length) + (n_images%terms.Length), /*out*/imageUrls);
			for (int i = 1; i < Math.Min(terms.Length, n_images); ++i)
				Search(searchservice, terms[i], Math.Max(1, n_images/terms.Length), /*out*/imageUrls);
			
			return imageUrls;
		}

		private void b(object sender, KeyPressEventArgs e)
		{
			if (ModifierKeys == Keys.Control)
			{	var font = BodyTextBox.SelectionFont;
				if (font.Bold) BodyTextBox.SelectionFont =
					new System.Drawing.Font(font, font.Style & ~System.Drawing.FontStyle.Bold);
				else BodyTextBox.SelectionFont =
					new System.Drawing.Font(font, font.Style |  System.Drawing.FontStyle.Bold);
			}
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
