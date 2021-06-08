using System;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using System.Collections.Generic;
using System.Windows.Forms;
//using System.Windows.Documents;
using Microsoft.Office.Core;
using Ppt = Microsoft.Office.Interop.PowerPoint;
using Google.Apis.Customsearch.v1;
using Google.Apis.Customsearch.v1.Data;
using ListRequest = Google.Apis.Customsearch.v1.CseResource.ListRequest;
using System.IO;
using System.Text;
using System.Drawing;

namespace SEHAmerica_ppt_Maker
{
	public partial class Form1 : Form
	{
		public const string save_err_msg = "Unexpected Error when trying to save to ";
		public const string wrIt_err_msg = "Unexpected Error when trying to write to ";
		public const string open_err_msg = "Unexpected Error when trying to open ";
		public const string load_err_msg = "Unexpected Error when trying to load ";
		public const string frOz_err_msg = "PowerPoint is in an unresponsive state.  Please resolve.";
		public const string auto_denied = "PowerPoint Helper was denied access to ";
		public const string powerpointformat = "Powerpoint (*.ppt;*.pptx;*.pptm)|*.ppt;*.pptx;*.pptm";
		
		Ppt.Application PowerPoint_App;
		Ppt.Presentations presentation_list;
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
		ImageList thumb_list, image_list;	//Lists of the "thumbnail" images and the full image links.
		OpenFileDialog ofd;
		List<string> box_list;
		readonly Point initloc;
		readonly Size change;				//How much the size of the ListView of images needs to change.
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
			image_list = new ImageList { ImageSize = new Size(160, 160) };
			box_list = new List<string>();
			initloc = ListView1.Location;
			change= new Size
				(ListView1.Width - BodyTextBox.Width, TitleBox.Location.Y - ListView1.Location.Y);

			PowerPoint_App = new Ppt.Application();
			presentation_list = PowerPoint_App.Presentations;
		}

		public bool NoPowerPoint()
		{	
			try
			{	if (presentation_list.Count < 1 || "PowerPoint" == PowerPoint_App.Caption)
					return true;
				else if (null == presentation)
				{	presentation = presentation_list[presentation_list.Count];
					ChangeButtons();
				}
				return false;	//Actually ("PowerPoint" == PowerPoint_App.Caption) and (presentation_list.Count == 1) should be equivilent.
			}
			catch(System.Runtime.InteropServices.COMException exc)
			{	MessageBox.Show(frOz_err_msg);
				return true;
			}
		}

		private bool CreateNewPresentation()
		{
			SaveFileDialog sfd = new SaveFileDialog { Filter = powerpointformat };
			try
			{	if (sfd.ShowDialog() == DialogResult.OK) file_path = sfd.FileName;
				else return false;
			} catch (SystemException exc)
			{	MessageBox.Show(save_err_msg + sfd.FileName + ":\n\n" + exc);
				return false;
			} finally { sfd.Dispose(); }

			presentation = presentation_list.Add();
			layout = Ppt.PpSlideLayout.ppLayoutText;
			presentation.Slides.AddSlide(1, presentation.SlideMaster.CustomLayouts[layout]);
			ChangeButtons();
			return true;
		}

		private bool LoadPresentation()
		{	
			ofd = new OpenFileDialog { Filter = powerpointformat };
			try
			{	if (ofd.ShowDialog() == DialogResult.OK) file_path = ofd.FileName;
				else return false;
			} catch (SystemException exc)
			{	MessageBox.Show(open_err_msg + ofd.FileName + ":\n\n" + exc);
				return false;
			} finally { ofd.Dispose(); }

			presentation = presentation_list.Open(file_path, msoFalse, msoFalse, msoTrue);
			ChangeButtons();
			return true;
		}
		
		private void ReadSlide(object sender, EventArgs e)
		{
			if (NoPowerPoint()) if(! LoadPresentation() ) return;//If there is no presentation open load one, but return if the user cancels.
			
			TitleBox.Clear();
			BodyTextBox.Clear();
			BodyTextBox.ZoomFactor = 1f;//This line is to alleviate a bug in the RichTextBox class that
			//^prevents display of the proper ZoomFactor after clearing it.

			foreach (var item in presentation.Slides[1].Shapes)
			{	var shape = (Ppt.Shape)item;
				if (shape.HasTextFrame != msoTrue) continue;

				if (shape.Name.Contains("Title"))
				{	TitleBox.Text += shape.TextFrame.TextRange.Text + ' ';
					box_list.Insert(0, shape.Name);
				}
				
				else if(shape.TextFrame.HasText == msoTrue)
				{	IDataObject clipped = Clipboard.GetDataObject();
					shape.TextFrame.TextRange.Copy();
					BodyTextBox.Paste();
					//if (BodyTextBox.Font.SizeInPoints > 17f) BodyTextBox.ZoomFactor = 0.5f;
					if (shape.TextFrame.TextRange.Font.Size > 17f) BodyTextBox.ZoomFactor = 0.5f;
					Clipboard.SetDataObject(clipped, true);
					box_list.Add(shape.Name);
				}
			}
		}

		/* While I really should take some of the conditions outside of the loop; seeing
		 * as I already know of a bug I'd need to fix inside If Else block, and I don't
		 * even know if the optimizer will do it for me; I decided I'd just make a note
		 * of it, and leave all of the stuff requiring repeating code for later.
		 */
		private void Save(object sender, EventArgs e)
		{	
			bool new_presentation = NoPowerPoint();
			try
			{	if (new_presentation) if (! CreateNewPresentation() ) return;//If there is no presentation open create one, but return if the user cancels.
				foreach (var item in presentation.Slides[1].Shapes)
				{	var shape = (Ppt.Shape)item;
					if (shape.HasTextFrame != msoTrue) continue;

					if ((box_list.Count > 0 && shape.Name == box_list[0]) ||
						(box_list.Count < 1 && shape.Name.Contains("Title")))
						shape.TextFrame.TextRange.Text = TitleBox.Text;
				
					else if(box_list.Contains(shape.Name) || presentation.Slides[1].Shapes.Count == 2)
					{	IDataObject clipped = Clipboard.GetDataObject();
						BodyTextBox.SelectAll();
						BodyTextBox.Copy();
						shape.TextFrame.TextRange.Paste();
						Clipboard.SetDataObject(clipped, true);
					}
				}

				if (new_presentation) try { presentation.SaveAs(file_path); }
					catch (FileNotFoundException)
						{ MessageBox.Show(save_err_msg + file_path + " (invalid name)."); }
					catch (System.Runtime.InteropServices.COMException exc)
						{ MessageBox.Show(save_err_msg + file_path + ":\n\n" + exc); }

			} catch (System.Runtime.InteropServices.COMException exc)
			{	if (exc.ErrorCode == -2147188160)
					 MessageBox.Show(auto_denied + (file_path?? "the presentation.")+ "\n\n" + exc);
				else MessageBox.Show(wrIt_err_msg + file_path +  ":\n\n" + exc);
			} catch (ArgumentException exc)
				{ MessageBox.Show(wrIt_err_msg + file_path + ":\n\n" + exc); }
		}

		private string[] ParseText(string titletext, string richtext)
		{	
			const StringComparison matchKs = StringComparison.Ordinal;
			List<string> terms = new List<string>(new string[] { titletext });
			List<int> startindices = new List<int>(), endindices = new List<int>{-3};
			//char[] whitespace = {' ','\n','\t'};
			
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
					starti = richtext.IndexOf(" ", starti, matchKs);
					if (-1 == starti) break;	//An empty bolded range of text with return characters for delimination can appear at the end of rich text.
					startindices.Add(starti);
					var endi = richtext.IndexOf("\\b0" , startindices[i] + 1, matchKs);
					if (-1 == endi) break;
					endindices.Add(endi);
				} 
			} catch(IndexOutOfRangeException) { }
			
			for (int i = 0; i < startindices.Count; ++i)
			{	int begin = startindices[i];
				var boldedtext = richtext.Substring(begin, endindices[i+1] - begin);
				var filteredtext = RemovePunct(boldedtext, '.', '\"','\'','?','*','\r').Trim();
				if (string.Empty != filteredtext) terms.Add(filteredtext);
			}

			return terms.ToArray();
		}

		private string[] ParseTextDirect()
		{	try
			{	const StringSplitOptions nOMTs = StringSplitOptions.RemoveEmptyEntries;
				List<string> terms = new List<string>();
				Ppt.TextRange run;

				foreach (var item in presentation.Slides[1].Shapes)
				{	var shape = (Ppt.Shape)item;
					if (shape.HasTextFrame != msoTrue) continue;

					if (shape.Name.Contains("Title"))
						terms.Insert(0, shape.TextFrame.TextRange.Text);
				
					else if(shape.TextFrame.HasText == msoTrue)
					{	for (int i = 1; i < image_count - 1; ++i)
						{	run = shape.TextFrame.TextRange.Runs(i);
							if (run == null) break;
							run.RemovePeriods();
							if (run.Font.Bold == msoTrue)
							{	var punctless = RemovePunct(run.Text.Trim(), '\"','\'','?','*','\r');
								terms.AddRange(punctless.Split(new char[]{'\t','\n'}, nOMTs));
							}
						}
					}
				}

				return terms.ToArray();
			} catch (System.Runtime.InteropServices.COMException exc)
			{	if (exc.ErrorCode == -2147188160)
					 MessageBox.Show(auto_denied + (file_path?? "the presentation.")+ "\n\n" + exc);
				else MessageBox.Show("Unexcpected error when trying to parse text." + "\n\n" + exc);
				return Array.Empty<string>();
			}
		}

		private void SearchImages(object sender, EventArgs e)
		{
			if (NoPowerPoint()) if(! LoadPresentation() ) return;//If there is no presentation open create one, but return if the user cancels.

			List<Double_Links> imgs;
			if (CheckNativeBoxes.Checked)
				 imgs = GoogleImageSearch(ParseText(TitleBox.Text, BodyTextBox.Rtf), image_count);
			else imgs = GoogleImageSearch(ParseTextDirect(), image_count);
			
			var pb = new PictureBox { SizeMode = PictureBoxSizeMode.Zoom };
			for(int i = 0; i < imgs.Count; ++i)
			{	try
				{	try { pb.Load(imgs[i].thumbnail); }
					catch (ArgumentException)
					{	try { pb.Load(imgs[i].picture); }
						catch (ArgumentException)
							{ pb.Image = pb.ErrorImage; }
					}
					thumb_list.Images.Add(pb.Image);
					try { pb.Load(imgs[i].picture); }
					catch (ArgumentException)
					{	try { pb.Load(imgs[i].thumbnail); }
						catch (ArgumentException)
							{ pb.Image = pb.ErrorImage; }
					}
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

		private void LoadImages(object sender, EventArgs e)
		{ 
			foreach (ListViewItem image in ListView1.SelectedItems)
			{	try { Active_slide.Shapes.AddPicture(image.Text, msoFalse, msoTrue, 0f, 0f); }
				catch (ArgumentException exc)
					{ MessageBox.Show(load_err_msg + image.Text + ":\n\n" + exc); }
			}
		}

		private void UseNativeTextBoxes(object sender, EventArgs e)
		{
			if (CheckNativeBoxes.Checked)
			{	label1.Show();
				TitleBox.Show();
				label2.Show();
				BodyTextBox.Show();
				BoldBtn.Show();
				//BoldBtn.Enabled = true;
				ListView1.Size += change;
				ListView1.Location = initloc;
				if (!NoPowerPoint())
				{	ImageBtn.Location = new Point(ImageBtn.Location.X, 156);
					//Read_Btn.Enabled = true;
					Read_Btn.Show();
					//WriteBtn.Enabled = true;
					WriteBtn.Show();
				}
			}
			else
			{	label1.Hide();
				TitleBox.Hide();
				label2.Hide();
				BodyTextBox.Hide();
				BoldBtn.Hide();
				//BoldBtn.Enabled = false;
				ListView1.Size = new Size
					(BodyTextBox.Width, ListView1.Height + ListView1.Location.Y - TitleBox.Location.Y);
				ListView1.Location = TitleBox.Location;
				if (!NoPowerPoint())
				{	ImageBtn.Location = Read_Btn.Location;
					Read_Btn.Hide();
					//Read_Btn.Enabled = false;
					WriteBtn.Hide();
					//WriteBtn.Enabled = false;
				}
			}
		}

		private void ChangeButtons()
		{
			Read_Btn.Text =  "Read";
			WriteBtn.Text = "Write";
			if (!CheckNativeBoxes.Checked)
			{	ImageBtn.Location = Read_Btn.Location;
				Read_Btn.Hide();
				//Read_Btn.Enabled = false;
				WriteBtn.Hide();
				//WriteBtn.Enabled = false;
			}
		}

		private void RichTextBox_KeyPress(object sender, KeyPressEventArgs e)
			{ if (e.KeyChar == '\u0002') ToggleBold(); }
		//	{ if (e.KeyChar == 'b' && ModifierKeys == Keys.Control) ToggleBold(); }

		private void BoldBtn_Click(object sender, EventArgs e) => ToggleBold();

		private void ToggleBold()
		{
			var font = BodyTextBox.SelectionFont;
			if (font.Bold)
			{	BodyTextBox.SelectionFont = new Font(font, font.Style & ~FontStyle.Bold);
				BoldBtn.BackColor = DefaultBackColor;
			}
			else /*!font.Bold*/
			{	BodyTextBox.SelectionFont = new Font(font, font.Style |  FontStyle.Bold);
				BoldBtn.BackColor = SystemColors.ActiveCaption;
			}
		}

		private void BodyTextBox_SelectionChanged(object sender, EventArgs e)
		{
			if (BodyTextBox.SelectionFont.Bold)
				 BoldBtn.BackColor = SystemColors.ActiveCaption;
			else BoldBtn.BackColor = DefaultBackColor;
		}

		private List<Double_Links> GoogleImageSearch(string[] terms, int n_images)
		{
			if (terms.Length == 0) return new List<Double_Links>();
			
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
