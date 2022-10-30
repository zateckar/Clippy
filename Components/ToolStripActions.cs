using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Clippy
{
    public partial class Main : Form
    {
        private void showInTaskbarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            settings.ShowInTaskBar = !settings.ShowInTaskBar;
            UpdateControlsStates();
        }

        private void showOnlyFavoriteToolStripMenuItem_Click(object sender = null, EventArgs e = null)
        {
            MarkFilter.SelectedValue = "favorite";
        }

        private void showOnlyFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TypeFilter.SelectedValue = "file";
        }

        private void showOnlyImagesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TypeFilter.SelectedValue = "img";
        }

        private void showOnlyTextsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TypeFilter.SelectedValue = "text";
        }

        private void showOnlyUsedToolStripMenuItem_Click(object sender = null, EventArgs e = null)
        {
            MarkFilter.SelectedValue = "used";
        }

        private void sortByCreationDateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sortField = "Clips.Created";
            ReloadList();
        }

        private void sortByDefaultToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sortField = "Clips.Id";
            ReloadList();
        }

        private void sortByVisualSizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //sortField = "Chars"; // not working
            sortField = "Clips.Chars";
            ReloadList();
        }


        private void toolStripApplicationCopyAll_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(textBoxApplication.Text))
                SetTextInClipboard(textBoxApplication.Text);
        }

        private void toolStripButtonAutoSelectMatch_Click(object sender, EventArgs e)
        {
            settings.AutoSelectMatch = !settings.AutoSelectMatch;
            UpdateControlsStates();
        }

        private void toolStripButtonFixedWidthFont_Click(object sender, EventArgs e)
        {
            settings.MonospacedFont = !settings.MonospacedFont;
            UpdateControlsStates();
            AfterRowLoad();
        }

        public void toolStripButtonMarkFavorite_Click(object sender, EventArgs e)
        {
            if (clip == null)
                return;

            bool state = BoolFieldValue("Favorite");

            bool stateneg = !BoolFieldValue("Favorite");

            SetRowMark("Favorite", !BoolFieldValue("Favorite"));
        }

        private void toolStripMenuItemMarkFavorite_Click(object sender, EventArgs e)
        {
            if (clip == null)
                return;

            bool state = BoolFieldValue("Favorite");

            bool stateneg = !BoolFieldValue("Favorite");

            SetRowMark("Favorite", !BoolFieldValue("Favorite"));
        }

        

        private void toolStripMenuItem20_Click(object sender, EventArgs e)
        {
            sortField = "Clips.Size";
            ReloadList();
        }



        private void toolStripMenuItemClearFilters_Click(object sender, EventArgs e)
        {
            ClearFilter();
            dataGridView.Focus();
        }

        private void toolStripMenuItemCompareLastClips_Click(object sender = null, EventArgs e = null)
        {
            if (lastClipWasMultiCaptured)
                notifyIcon.ShowBalloonTip(2000, Application.ProductName, Properties.Resources.LastClipWasMultiCaptured, ToolTipIcon.Info);
            //string sql = "SELECT Id FROM Clips ORDER BY Id Desc Limit 2";
            //SQLiteCommand commandSelect = new SQLiteCommand(sql, m_dbConnection);

            //using (SQLiteDataReader reader = commandSelect.ExecuteReader())
            //{
            //    int id1 = 0, id2 = 0;
            //    if (reader.Read())
            //        id1 = (int)reader["Id"];
            //    if (reader.Read())
            //        id2 = (int)reader["Id"];
            //    if (id2 > 0 && id1 > 0)
            //        CompareClipsByID(id2, id1);
            //}

            var col = liteDb.GetCollection<Clip>("clips").Query()
                            .OrderByDescending(x => x.Id)
                            .Select(x => x.Id)
                            .Limit(2)
                            .ToList();
            if (col.Count > 0)
            {
                int id1 = 0, id2 = 0;

                id1 = col[0];
                id2 = col[1];

                if (id2 > 0 && id1 > 0)
                    CompareClipsByID(id2, id1);

            }
        }

        private void toolStripMenuItemPasteCharsFast_Click(object sender, EventArgs e)
        {
            SendPasteOfSelectedTextOrSelectedClips(PasteMethod.SendCharsFast);
        }

        private void toolStripMenuItemPasteCharsSlow_Click(object sender, EventArgs e)
        {
            SendPasteOfSelectedTextOrSelectedClips(PasteMethod.SendCharsSlow);
        }

        private void toolStripMenuItemSecondaryColumns_Click(object sender = null, EventArgs e = null)
        {
            settings.ShowSecondaryColumns = !settings.ShowSecondaryColumns;
            UpdateControlsStates();
            UpdateColumnsSet();
        }

        private void toolStripMenuItemShowAllTypes_Click(object sender, EventArgs e)
        {
            TypeFilter.SelectedValue = "allTypes";
        }

        private void toolStripUrlCopyAll_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(urlTextBox.Text))
                SetTextInClipboard(urlTextBox.Text);
        }

        private void toolStripWindowTitleCopyAll_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(textBoxWindow.Text))
                SetTextInClipboard(textBoxWindow.Text);
        }

        private void textCompareToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string type1;
            int id1, id2;
            if (dataGridView.SelectedRows.Count == 0)
                return;
            if (dataGridView.SelectedRows.Count == 1)
            {
                type1 = clip.Type;
                if (!IsTextType() && type1 != "file")
                    return;
                if (_lastSelectedForCompareId == 0)
                {
                    _lastSelectedForCompareId = clip.Id;

                    return;
                }
                else
                {
                    id1 = clip.Id;
                    id2 = _lastSelectedForCompareId;
                    _lastSelectedForCompareId = 0;
                }
            }
            else
            {
                DataRowView row1 = (DataRowView)dataGridView.SelectedRows[0].DataBoundItem;
                id1 = (int)row1["id"];
                DataRowView row2 = (DataRowView)dataGridView.SelectedRows[1].DataBoundItem;
                id2 = (int)row2["id"];
            }
            CompareClipsByID(id2, id1);
        }

        private void setSelectedTextInFilterToolStripMenu_Click(object sender, EventArgs e)
        {
            AllowFilterProcessing = false;
            comboBoxSearchString.Text = richTextBox.SelectedText;
            AllowFilterProcessing = true;
            SearchStringApply();
        }

        private void setFavoriteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetRowMark("Favorite", true, true);
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            htmlTextBox.Document.ExecCommand("SelectAll", false, null);
        }
        private void moveTopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MoveSelectedRows(0);
        }

        private void moveUpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MoveSelectedRows(-1);
        }

        private void moveClipToTopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            settings.MoveCopiedClipToTop = !settings.MoveCopiedClipToTop;
            UpdateControlsStates();
        }

 

        private void moveDownToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MoveSelectedRows(1);
        }

        private void meandsAnySequenceOfCharsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            settings.SearchWildcards = !settings.SearchWildcards;
            UpdateControlsStates();
            SearchStringApply();
        }

        private void menuItemSetFocusClipText_Click(object sender, EventArgs e)
        {
            FocusClipText();
        }

        private void mergeTextOfClipsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateSelectedClipsHistory();
            string agregateTextToPaste = JoinOrPasteTextOfClips(PasteMethod.Null, out _);
            AddClip(null, null, "", "", "text", agregateTextToPaste);
            GotoLastRow(true);
        }

        private void MonitoringClipboardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SwitchMonitoringClipboard();
        }

        private void ChangeClipTitleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView.CurrentRow != null)
            {
                string oldTitle = clip.Title;

                TextBox tb = new()
                {
                    Size = new(splitContainer1.Panel1.Width - 10, 20),
                    Text = oldTitle
                };

                Button btnOK = new()
                {
                    Size = new(100, 30),
                    Text = "Save",
                    Location = new(splitContainer1.Panel1.Width - 250, 50),
                };
                //btnOK.Click += (sender, e) => BtnOK_Click(sender, e, tb);
                btnOK.Click += (sender, e) => ChangeClipTitleBtnOK_Click(tb);

                Button btnCancel = new()
                {
                    Size = new(100, 30),
                    Text = "Cancel",
                    Location = new(splitContainer1.Panel1.Width - 100, 50)
                };
                btnCancel.Click += ChangeClipTitleBtnCancel_Click;

                dataGridView.Visible = false;

                this.splitContainer1.Panel1.Controls.Add(tb);
                this.splitContainer1.Panel1.Controls.Add(btnOK);
                this.splitContainer1.Panel1.Controls.Add(btnCancel);


                //InputBoxResult inputResult = InputBox.Show(Properties.Resources.HowUseAutoTitle, Properties.Resources.EditClipTitle, oldTitle, this);


                //if (inputResult.ReturnCode == DialogResult.OK)
                //{
                //    //string sql = "Update Clips set Title=@Title where Id=@Id";
                //    //SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                //    //command.Parameters.AddWithValue("@Id", clip.Id);
                //    //string newTitle;
                //    //if (inputResult.Text == "")
                //    //    newTitle = TextClipTitle(clip.Text);
                //    //else
                //    //    newTitle = inputResult.Text;
                //    //command.Parameters.AddWithValue("@Title", newTitle);
                //    //command.ExecuteNonQuery();

                //    var clip_upd = liteDb.GetCollection<Clip>("clips").FindById(clip.Id);
                //    if (clip_upd != null)
                //    {

                //        if (inputResult.Text == "")
                //            clip_upd.Title = TextClipTitle(clip.Text);
                //        else
                //            clip_upd.Title = inputResult.Text;

                //        var col = liteDb.GetCollection<Clip>("clips").Update(clip_upd);
                //    }

                //    ReloadList(true);
                //}
            }
        }

        private void ChangeClipTitleBtnCancel_Click(object sender, EventArgs e)
        {
            dataGridView.Visible = true;
        }

        private void ChangeClipTitleBtnOK_Click(TextBox tb)
        {
            var clip_upd = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Id == clip.Id);
            if (clip_upd != null)
            {
                if (tb.Text == "")
                    clip_upd.Title = TextClipTitle(clip.Text);
                else
                    clip_upd.Title = tb.Text;

                _ = liteDb.GetCollection<Clip>("clips").Update(clip_upd);
            }
            dataGridView.Visible = true;
            ReloadList(true);
        }


        private void IgnoreBigTextClipsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            settings.SearchIgnoreBigTexts = !settings.SearchIgnoreBigTexts;
            UpdateControlsStates();
            SearchStringApply();
        }

        private void GotoTopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GotoLastRow();
        }

        

        private void everyWordIndependentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            settings.SearchWordsIndependently = !settings.SearchWordsIndependently;
            UpdateControlsStates();
            SearchStringApply();
        }

        private void exitToolStripMenuItem_Click(object sender = null, EventArgs e = null)
        {
            Application.Exit();
        }

        private void exitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            exitToolStripMenuItem_Click();
        }

        private void deleteAllNonFavoriteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(this, Properties.Resources.ConfirmDeleteAllNonFavorite, Properties.Resources.Confirmation, MessageBoxButtons.OKCancel);
            if (result != DialogResult.OK)
                return;
            deleteAllNonFavoriteClips();
            ReloadList();
        }

        private void decodeTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateSelectedClipsHistory();
            string selectedText = getSelectedOrAllText();
            string newSelectedText = WebUtility.UrlDecode(selectedText);
            if (newSelectedText.Equals(selectedText))
                return;
            AddClip(null, null, "", "", "text", newSelectedText);
            GotoLastRow(true);
        }

        private void copyClipToolStripMenuItem_Click(object sender = null, EventArgs e = null)
        {
            CopyClipToClipboard(null, false, settings.MoveCopiedClipToTop);
        }

        private void copyFullFilenameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (clip == null)
                return;
            string fullFilename = clip.AppPath.ToString();
            SetTextInClipboard(fullFilename);
        }

        private void copyLinkAdressToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string href = lastClickedHtmlElement.GetAttribute("href");
            SetTextInClipboard(href);
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataObject dto = new DataObject();
            string text = richTextBox.SelectedText;
            SetTextInClipboardDataObject(dto, text);
            if (settings.ShowNativeTextFormatting)
            {
                dto.SetText(richTextBox.SelectedRtf, TextDataFormat.Rtf);
            }
            SetClipboardDataObject(dto);
        }

        private void caseSensetiveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            settings.SearchCaseSensitive = !settings.SearchCaseSensitive;
            UpdateControlsStates();
            SearchStringApply();
        }

        private void clearClipboardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
        }

        private void autoselectFirstMatchedClipToolStripMenuItem_Click(object sender, EventArgs e)
        {
            settings.AutoSelectMatchedClip = !settings.AutoSelectMatchedClip;
            UpdateControlsStates();
        }

        private void filterByDateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            monthCalendar1.Show();
            monthCalendar1.Focus();
        }

        private void filterListBySearchStringToolStripMenuItem_Click(object sender, EventArgs e)
        {
            settings.FilterListBySearchString = !settings.FilterListBySearchString;
            UpdateControlsStates();
            if (!settings.FilterListBySearchString)
            {
                ReloadList();
                UpdateSearchMatchedIDs();
            }
            else
            {
                SearchStringApply();
            }
        }

        private void fitFromInsideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImageControl.ZoomFitInside();
        }

        private void openFavoritesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForPaste(true);
        }

        private void openInDefaultApplicationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenClipFile();
        }

        private void openLastClipsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForPaste(false, true);
        }

        private void openLinkInBrowserToolStripMenuItem_Click(object sender = null, EventArgs e = null)
        {
            string href = lastClickedHtmlElement.GetAttribute("href");
            int textOffset = -1;

            if (textOffset >= 0 && OpenLinkFromTextBox(TextLinkMatches, textOffset, false))
            {
            }
            else
            {
                if (!String.IsNullOrEmpty(href))
                    Process.Start(href);
            }
        }

        private void openWindowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForPaste(false, false, true);
        }

        private void openWithToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenClipFile(false);
        }

        private void originalSizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImageControl.ZoomOriginalSize();
        }

        private void pasteAsTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SendPasteOfSelectedTextOrSelectedClips(PasteMethod.Text);
        }

        private void pasteFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SendPasteOfSelectedTextOrSelectedClips(PasteMethod.File);
        }

        private void pasteIntoSearchFieldMenuItem_Click_1(object sender, EventArgs e)
        {
            string selectedText = "";
            string agregateTextToPaste = GetSelectedTextOfClips(ref selectedText);
            agregateTextToPaste = ConvertTextToLine(agregateTextToPaste);
            comboBoxSearchString.Text = agregateTextToPaste.Substring(0, Math.Min(50, agregateTextToPaste.Length));
        }

        private void pasteLineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SendPasteOfSelectedTextOrSelectedClips(PasteMethod.Line);
        }

        private void pasteOriginalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SendPasteOfSelectedTextOrSelectedClips();
        }

        private void returnToPrevousSelectedClipToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int newSelectedID = 0;
            if (selectedClipsBeforeFilterApply.Count > 0)
            {
                newSelectedID = selectedClipsBeforeFilterApply[selectedClipsBeforeFilterApply.Count - 1];
                selectedClipsBeforeFilterApply.RemoveAt(selectedClipsBeforeFilterApply.Count - 1);
            }
            ClearFilter(newSelectedID);
            dataGridView.Focus();
        }

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PropertyGrid pg = new PropertyGrid();


            if (settingsToolStripMenuItem1.BackColor != Color.LightGray)
            {
                dataGridView.Visible = false;

                
                pg.PropertyValueChanged += new PropertyValueChangedEventHandler(propertyGrid_PropertyValueChanged);
                splitContainer2.Panel1.Controls.Add(pg);
                pg.Dock = DockStyle.Fill;
                pg.HelpVisible = false;
                pg.ToolbarVisible = false;
                pg.PropertySort = PropertySort.Alphabetical;
                pg.SelectedObject = settings;
                settingsToolStripMenuItem1.BackColor = Color.LightGray;
            }
            else
            {
                splitContainer2.Panel1.Controls.Remove(pg);
                dataGridView.Visible = true;
                settingsToolStripMenuItem1.BackColor = DefaultBackColor;
            }

        }

        private void showAllMarksToolStripMenuItem_Click(object sender = null, EventArgs e = null)
        {
            MarkFilter.SelectedValue = "allMarks";
        }

        private void resetFavoriteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetRowMark("Favorite", false, true);
        }


        private void toolStripButtonWordWrap1_Click(object sender, EventArgs e)
        {
            settings.WordWrap = !settings.WordWrap;
            allowTextPositionChangeUpdate = false;
            UpdateControlsStates();
            allowTextPositionChangeUpdate = true;
            if (clip.Type == "html")
                AfterRowLoad();
            OnClipContentSelectionChange();
        }


        private void toolStripButtonSecondaryColumns1_Click(object sender, EventArgs e)
        {
            settings.ShowSecondaryColumns = !settings.ShowSecondaryColumns;
            UpdateControlsStates();
            UpdateColumnsSet();
        }

        private void toolStripButtonTextFormatting1_Click(object sender, EventArgs e)
        {
            settings.ShowNativeTextFormatting = !settings.ShowNativeTextFormatting;
            UpdateControlsStates();
            string clipType = clip.Type;
            if (clipType == "html" || clipType == "rtf")
                AfterRowLoad(true);
        }

        private void toolStripButtonMonospacedFont1_Click(object sender, EventArgs e)
        {
            settings.MonospacedFont = !settings.MonospacedFont;
            UpdateControlsStates();
            AfterRowLoad();
        }


        public void toolStripMenuItemEditClipText1_Click(object sender = null, EventArgs e = null)
        {
            if (clip == null)
                return;
            bool newEditMode = !EditMode;
            string clipType = clip.Type.ToString();
            if (!IsTextType())
                return;

            allowRowLoad = false;
            if (!newEditMode)
            {
                ReloadList();
            }
            else
            {
                if (clipType != "text")
                {
                    AddClip(null, null, "", "","text", clip.Text, clip.Application,clip.Window, "", 0, clip.AppPath, clip.Used, false,true,clip.Title);
                    //AddClip(null, null, "", "", "text", clip.Text);
                    GotoLastRow(true);
                }
                //else
                //    UpdateClipBindingSource();
            }
            allowRowLoad = true;
            EditMode = newEditMode;
            AfterRowLoad(true, -1);
           UpdateControlsStates();
        }

        private void moveCopiedClipToTopToolStripButton1_Click(object sender, EventArgs e)
        {
            settings.MoveCopiedClipToTop = !settings.MoveCopiedClipToTop;
            UpdateControlsStates();
        }

        private void toolStripButtonClearFilterAndSelectTop1_Click(object sender, EventArgs e)
        {
            ClearFilter(-1);
            dataGridView.Focus();
        }

        private void toolStripButtonTopMostWindow1_Click(object sender, EventArgs e)
        {
            this.TopMost = !this.TopMost;
            windowAlwaysOnTopToolStripMenuItem.Checked = this.TopMost;
            //toolStripButtonTopMostWindow.Checked = this.TopMost;

            if (this.TopMost)
                toolStripButtonTopMostWindow1.BackColor = Color.LightBlue;
            else
                toolStripButtonTopMostWindow1.BackColor = Control.DefaultBackColor;
        }


    }
}
