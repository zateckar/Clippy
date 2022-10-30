using LiteDB;
using Microsoft.Win32;
using mshtml;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Cursor = System.Windows.Forms.Cursor;
using KeyEventArgs = System.Windows.Forms.KeyEventArgs;
using MouseEventArgs = System.Windows.Forms.MouseEventArgs;
using MouseEventHandler = System.Windows.Forms.MouseEventHandler;
using Timer = System.Windows.Forms.Timer;

namespace Clippy
{
     public partial class Main : Form
    {
        public const string IsMainPropName = "IsMain";
        public ResourceManager CurrentLangResourceManager;
        public KeyboardHook keyboardHook;
        //public string Locale = "";

        private WinEventDelegate dele = null;
        private IntPtr HookChangeActiveWindow;
        private static RECT lastChildWindowRect;
        private static IntPtr lastActiveParentWindow;
        private static IntPtr lastChildWindow;
        private static Point lastCaretPoint;

        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);
        public delegate void MouseMovedEvent();

        private const uint EVENT_SYSTEM_FOREGROUND = 3;
        //private const int MM_ANISOTROPIC = 8;
        //private const int MM_ISOTROPIC = 7;
        private const int SC_MINIMIZE = 0xf020;
        private const int tabLength = 4;
        private const uint WINEVENT_OUTOFCONTEXT = 0;
        private const int WM_SYSCOMMAND = 0x0112;
        //private const int WS_EX_COMPOSITED = 0x02000000;
        //private const int WS_EX_NOACTIVATE = 0x08000000;

        private static Dictionary<string, string> TextPatterns = new()
        {
            {"time", "((" + datePattern + "\\s" + timePattern + ")|(?:(" + timePattern + "\\s)?" + datePattern + ")|(?:" + timePattern + "))"},
            {"email", "(\\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,6}\\b)"},
            {"number", "((?:(?:\\s|^)[-])?\\b[0-9]+\\.?[0-9]+)\\b"},
            {"phone", "(?:[\\s\\(]|^)(\\+?\\b\\d?(\\d[ \\-\\(\\)]{0,2}){7,19}\\b)"},
            {"url", "(\\b(?:https?|ftp|file)://[-A-Z0-9+&@#\\\\/%?=~_|!:,.;]*[A-Z0-9+&@#/%=~_|])"},
            {"url_image",  @"(https?:\/\/.*\.(?:png|jpg|gif|jpeg|svg))"},
            {"url_video",  @"(?:https?://)?(?:www.)?youtu(?:\.be|be\.com)/(?:(?:.*)v(?:/|=)|(?:.*/)?)([a-zA-Z0-9-_]+)(?:[&?][%a-zA-Z0-9-_]+=[%a-zA-Z0-9-_]+)*"},
            {"filename", @"((?:\b[a-z]:|\\\\[a-z0-9 %._-]+\\[a-z0-9 $%._-]+)\\(?:[^\\/:*?""<>|\r\n]+\\)*[^\\/:*?""<>|\r\n]*)"}
        };
        private static Dictionary<string, Bitmap> brightIconCache = new();
        private static Dictionary<string, Bitmap> originalIconCache = new();
        //private static Dictionary<string, object> clipboardContents = new Dictionary<string, object>();
        private static List<LastClip> lastClips = new();
        private Dictionary<int, DateTime> lastPastedClips = new();

        private static Color _usedColor = Color.FromArgb(210, 255, 255);
        private static Color favoriteColor = Color.FromArgb(255, 230, 220);
        private Color[] _wordColors = new Color[] { Color.Red, Color.DeepPink, Color.DarkOrange };

        //private const int ChannelDataLifeTime = 60;
        private const int maxClipsToSelect = 300;
        private const int MaxTextViewSize = 10000;
        private const int ClipTitleLength = 70;

        private static readonly string LinkPattern = TextPatterns["url"];
        private static readonly string fileOrFolderPattern = TextPatterns["filename"];
        private static readonly string imagePattern = TextPatterns["url_image"];
        private static readonly string videoPattern = TextPatterns["url_video"];
        private static readonly string datePattern = "\\b(?:19|20)?[0-9]{2}[\\-/.][0-9]{2}[\\-/.](?:19|20)?[0-9]{2}\\b";
        private static readonly string timePattern = "\\b[012]?\\d:[0-5]?\\d(?::[0-5]?\\d)?\\b";
        
        private readonly RichTextBox _richTextBox = new();
        private readonly RichTextBox richTextInternal = new();
        private int _lastSelectedForCompareId;
        private int clipRichTextLength = 0;
        private int factualTop = 0;
        private int maxWindowCoordForHiddenState = -10000;
        private int LastId = 0;
        private int selectedRangeStart = -1;
        private int selectionLength = 0;
        private int selectionStart = 0;
        private bool EditMode = false;
        private bool AllowFilterProcessing = false;
        private bool AllowHotkeyProcess = true;
        private bool allowProcessDataGridSelectionChanged = true;
        private bool allowRowLoad = true;
        private bool allowTextPositionChangeUpdate = false;
        private bool allowVisible = false;
        private bool areDeletedClips = false;
        private bool filterOn = false;
        private bool periodFilterOn = false;
        private bool lastClipWasMultiCaptured = false;
        private bool htmlInitialized = false;
        private bool htmlMode = false;
        private bool TextWasCut;
        private bool titleToolTipShown = false;
        private bool UsualClipboardMode = false;
        private readonly string DataFormat_ClipboardViewerIgnore = "Clipboard Viewer Ignore";
        private readonly string DataFormat_RemoveTempClipsFromHistory = "RemoveTempClipsFromHistory";
        private readonly string DataFormat_XMLSpreadSheet = "XML SpreadSheet";
        private string sortField = "Id";
        private string searchString = "";

        public Clip clip;
        public Settings settings;
        public LiteDatabase liteDb;
        //public LiteDatabase defaultDb;
        public bool oneDB = true;

        private StringCollection ignoreModulesInClipCapture;
        private Bitmap imageFile;
        private Bitmap imageHtml;
        private Bitmap imageImg;
        private Bitmap imageRtf;
        private Bitmap imageText;
        private HtmlElement lastClickedHtmlElement;

        private Point lastMousePoint;
        private Point LastMousePoint;
        private Point PreviousCursor;
        private DateTime lastPasteMoment = DateTime.Now;
        private DateTime lastReloadListTime;
        private DateTime lastCaptureMoment = DateTime.Now;
        private DateTime TimeFromWindowOpen;

        private Task<DataTable> lastReloadListTask;
        private List<int> searchMatchedIDs = new();
        private List<int> selectedClipsBeforeFilterApply = new();

        private Timer tempCaptureTimer = new();
        private Timer captureTimer = new();
        private Timer titleToolTipBeforeTimer = new();
        private ToolTip titleToolTip = new();
        private MatchCollection TextLinkMatches;
        private MatchCollection FilterMatches;
        private MatchCollection UrlLinkMatches;

        public Main(bool StartMinimized)
        {
            //this.UserSettingsPath = UserSettingsPath;
            //this.PortableMode = PortableMode;

            InitializeComponent();

            toolStripSearchOptions.DropDownDirection = ToolStripDropDownDirection.AboveRight;
            dele = new WinEventDelegate(WinEventProc);
            HookChangeActiveWindow = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, dele, 0, 0, WINEVENT_OUTOFCONTEXT);
            // register the event that is fired after the key press.
            keyboardHook = new KeyboardHook(CurrentLangResourceManager);
            keyboardHook.KeyPressed += new EventHandler<KeyPressedEventArgs>(hook_KeyPressed);
            ImageControl.MouseWheel += new MouseEventHandler(ImageControl_MouseWheel);
            captureTimer.Tick += delegate { CaptureClipboardData(); };
            
            //toolStripTop.Renderer = new MyToolStripRenderer();
            toolStripBottom.Renderer = new MyToolStripRenderer();

            ResourceManager resourceManager = Properties.Resources.ResourceManager;
            imageText = resourceManager.GetObject("TypeText") as Bitmap;
            imageHtml = resourceManager.GetObject("TypeHtml") as Bitmap;
            imageRtf = resourceManager.GetObject("TypeRtf") as Bitmap;
            imageFile = resourceManager.GetObject("TypeFile") as Bitmap;
            imageImg = resourceManager.GetObject("TypeImg") as Bitmap;

            htmlTextBox.Navigate("about:blank");
            htmlTextBox.Document.ExecCommand("EditMode", false, null);

            PropertyInfo aProp = typeof(Control).GetProperty("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance);
            aProp.SetValue(dataGridView, true, null);

            BindingList<ListItemNameText> _comboItemsTypes = new()
            {
                new ListItemNameText {Name = "allTypes"},
                new ListItemNameText {Name = "img"},
                new ListItemNameText {Name = "file"},
                new ListItemNameText {Name = "text"},
                new ListItemNameText {Name = "rtf"},
                new ListItemNameText {Name = "html"},
            };
            foreach (KeyValuePair<string, string> pair in TextPatterns)
            {
                _comboItemsTypes.Add(new ListItemNameText { Name = "text_" + pair.Key });
            }
            TypeFilter.DataSource = _comboItemsTypes;
            TypeFilter.DisplayMember = "Text";
            TypeFilter.ValueMember = "Name";
            TypeFilter.SelectedIndex = 0;

            BindingList<ListItemNameText> _comboItemsMarks = new()
            {
                new ListItemNameText { Name = "allMarks", Text = "allMarks" },
                new ListItemNameText { Name = "favorite", Text = "favorite" },
                new ListItemNameText { Name = "used", Text = "used" }
            };
            MarkFilter.DataSource = _comboItemsMarks;
            MarkFilter.DisplayMember = "Text";
            MarkFilter.ValueMember = "Name";
            MarkFilter.SelectedIndex = 0;
            

            (dataGridView.Columns["AppImage"] as DataGridViewImageColumn).DefaultCellStyle.NullValue = null;
            richTextBox.AutoWordSelection = false;
            urlTextBox.AutoWordSelection = false;

            //if (!Directory.Exists(UserSettingsPath))
            //{
            //    Directory.CreateDirectory(UserSettingsPath);
            //}
            OpenDatabase();
            RegisterHotKeys();
            FillFilterItems();
            ConnectClipboard();

            // Initialize StringCollection settings to prevent error saving settings
            if (settings.LastFilterValues == null)
            {
                settings.LastFilterValues = new List<string>();
            }
            if (settings.IgnoreApplicationsClipCapture == null)
            {
                settings.IgnoreApplicationsClipCapture = new List<string>();
            }

            if (true && (settings.WindowPositionX != 0 || settings.WindowPositionY != 0) && (settings.WindowPositionX != -32000 && settings.WindowPositionY != -32000) // old version could save minimized state coords
                    && settings.WindowPositionY != maxWindowCoordForHiddenState )
            {
                this.Left = settings.WindowPositionX;
                this.Top = settings.WindowPositionY;
            }

            if (StartMinimized == false)
                allowVisible = true;

            if (settings.MainWindowSizeWidth > 0)
            {
                this.Size = new Size(settings.MainWindowSizeWidth, settings.MainWindowSizeHeight);
            }
                
            timerReconnect.Interval = (1000 * 5); // 5 seconds
            timerReconnect.Start();
            this.ActiveControl = dataGridView;
            ResetIsMainProperty();
            LoadSettings();

            if (settings.dataGridViewWidth != 0)
                splitContainer1.SplitterDistance = settings.dataGridViewWidth;
        }

        private void Main_Load(object sender, EventArgs e)
        {
            // Due to the hidden start window can be shown and this event raised not on the start
            // So we do not use it and make everything in constructor
            UpdateControlsStates(); //
            RestoreWindowIfMinimized();
            dataGridView.Focus();
        }

        private void OpenDatabase()
        {
            string DatabasePath = Application.StartupPath + "clippy.litedb";

            RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Clippy", true);
            if (key != null)
            {
                try
                {
                    if (!String.IsNullOrEmpty(key.GetValue("dbPath").ToString()))
                    {
                        DatabasePath = key.GetValue("dbPath").ToString();
                    }
                }
                catch { }

            }
            else
            {
                RegistryKey keysCreate = Registry.CurrentUser.CreateSubKey("SOFTWARE\\Clippy",true);
                keysCreate.SetValue("dbPath", DatabasePath);
            }

            if (!File.Exists(DatabasePath))
            {
                using (liteDb = new LiteDatabase(DatabasePath))
                {
                    var set = liteDb.GetCollection<Settings>("settings").Insert(new Settings());

                    var clip = new Clip
                    {
                        Type = "text",
                        Text = "My first clip",
                        Title = "My first clip",
                        Hash = "PC1BCXQEdgit/zDfA0KZgA==",
                        Chars = 100,
                        Id = 1,
                        Created = DateTime.Now,
                        Binary = null,
                        RichText = null,
                        HtmlText = null,
                        Used = false,
                        Favorite = false,
                        ImageSample = null,
                        Size = 100,
                    };

                    var col = liteDb.GetCollection<Clip>("clips");
                    col.EnsureIndex(x => x.Hash);
                    col.Insert(clip);
                }
            }

            liteDb = new LiteDatabase("Filename="+ DatabasePath +";Connection=direct");
            settings = liteDb.GetCollection<Settings>("settings").FindById(1);
            settings.DatabasePath = DatabasePath;
            settings.DatabaseSize = (new FileInfo(DatabasePath).Length / 1024).ToString() + " kB";

            dataGridView.DataSource = clipBindingSource;
            ReloadList();
            AllowFilterProcessing = true;
           
        }

        private Task<DataTable> ReloadListAsync(string TypeFilterSelectedValue, string MarkFilterSelectedValue, DateTime monthCalendar1SelectionStart, DateTime monthCalendar1SelectionEnd)
        {
            
            allowRowLoad = false;
            bool oldFilterOn = filterOn;
            filterOn = false;
            string filterValue = "";

            //liteDb = new LiteDatabase(DbFileName);
            //var colAll = liteDb.GetCollection<Clip>("clips").FindAll();

            var col = liteDb.GetCollection<Clip>("clips").Query()
                .Select(x => new { x.Id, x.Used, x.Title, x.Chars, x.Type, x.Favorite, x.ImageSample, x.AppPath, x.Size, x.Created })
                .ToEnumerable();

            if (!String.IsNullOrEmpty(searchString) && settings.FilterListBySearchString)
            {
                //sqlFilter += SqlSearchFilter();
                filterOn = true;
                if (settings.SearchIgnoreBigTexts)
                {
                    //sqlFilter += " AND (Chars < 100000 OR type = 'img')";
                    col = col.Where(w => w.Chars < 10000 || w.Type == "img");
                }
            }
            if (TypeFilterSelectedValue as string != "allTypes")
            {
                filterValue = TypeFilterSelectedValue as string;
                bool isText = filterValue.Contains("text");
                string filterValueText;
                if (isText)
                    filterValueText = "'html','rtf','text'";
                else
                    filterValueText = "'" + filterValue + "'";
                //sqlFilter += " AND type IN (" + filterValueText + ")";
                if (isText && filterValue != "text")
                {
                    //sqlFilter += String.Format(" AND Contain_{0}", filterValue.Substring("text_".Length));
                }
                filterOn = true;
            }
            if (MarkFilterSelectedValue as string == "favorite")
            {
                //filterValue = MarkFilterSelectedValue as string;
                //sqlFilter += " AND " + filterValue;
                col = col.Where(w => w.Favorite);
                filterOn = true;
            }
            if (MarkFilterSelectedValue as string == "fused")
            {
                col = col.Where(w => w.Used);
                filterOn = true;
            }
            if (periodFilterOn)
            {
                //sqlFilter += " AND Created BETWEEN @startDate AND @endDate ";
                //backgroundDataAdapter.SelectCommand.Parameters.AddWithValue("startDate", monthCalendar1SelectionStart);
                //backgroundDataAdapter.SelectCommand.Parameters.AddWithValue("endDate", monthCalendar1SelectionEnd);

                col = col.Where(w => w.Created > monthCalendar1SelectionStart && w.Created < monthCalendar1SelectionEnd);
                filterOn = true;
            }
            if (!oldFilterOn && filterOn)
                selectedClipsBeforeFilterApply.Clear();

            //string selectCommandText = "Select Id, NULL AS Used, NULL AS Title, NULL AS Chars, NULL AS Type, NULL AS Favorite, NULL AS ImageSample, NULL AS AppPath, NULL AS Size, NULL AS Created From Clips";
            //selectCommandText += " WHERE " + sqlFilter;
            //selectCommandText += " ORDER BY " + sortField + " desc";
            //if (settings.SearchCaseSensitive)
            //    selectCommandText = "PRAGMA case_sensitive_like = 1; " + selectCommandText;
            //else
            //    selectCommandText = "PRAGMA case_sensitive_like = 0; " + selectCommandText;
            //backgroundDataAdapter.SelectCommand.CommandText = selectCommandText;

            //DataTable table = new DataTable();
            //table.Locale = CultureInfo.InvariantCulture;
            //try
            //{
            //    backgroundDataAdapter.Fill(table);
            //}
            //catch (Exception)
            //{
            //    //int dummy = 0;
            //}

            col = col.OrderByDescending(w => w.Id);

            DataTable table1 = new DataTable();
            col.Fill(ref table1);

            settings.NumberOfClips = col.ToList().Count.ToString();

            return Task.FromResult(table1);
        }

        private bool AddClip(byte[] binaryBuffer = null, byte[] imageSampleBuffer = null, string htmlText = "", string richText = "", string typeText = "text", string plainText = "", string applicationText = "", string windowText = "", string url = "", int chars = 0, string appPath = "", bool used = false, bool favorite = false, bool updateList = true, string clipTitle = "", DateTime created = new DateTime())
        {
            DateTime dtNow = DateTime.Now;
            int msFromLastCapture = DateDiffMilliseconds(lastCaptureMoment);
            if (plainText == null)
                plainText = "";
            if (richText == null)
                richText = "";
            if (htmlText == null)
                htmlText = "";
            int byteSize = 0;
            CalculateByteAndCharSizeOfClip(htmlText, richText, plainText, ref chars, ref byteSize);
            if (binaryBuffer != null)
                byteSize += binaryBuffer.Length;
            if (byteSize > settings.MaxClipSizeKB * 1000)
            {
                string message = String.Format(Properties.Resources.ClipWasNotCaptured, (int)(byteSize / 1024), settings.MaxClipSizeKB, typeText);
                notifyIcon.ShowBalloonTip(2000, Application.ProductName, message, ToolTipIcon.Info);
                return false;
            }

            int oldCurrentClipId = 0;
            lastClipWasMultiCaptured = false;
            if (DateTime.MinValue == created)
                created = DateTime.Now;
            if (String.IsNullOrEmpty(clipTitle))
                clipTitle = TextClipTitle(plainText);
            string hash;
            //string sql = "SELECT Id, Title, Used, Favorite, Created FROM Clips Where Hash = @Hash";
            if (settings.ReplaceDuplicates)
            {
                MD5 md5 = MD5.Create();
                if (binaryBuffer != null)
                    md5.TransformBlock(binaryBuffer, 0, binaryBuffer.Length, binaryBuffer, 0);
                byte[] binaryText = Encoding.Unicode.GetBytes(plainText);
                md5.TransformBlock(binaryText, 0, binaryText.Length, binaryText, 0);
                if (settings.UseFormattingInDuplicateDetection || String.IsNullOrEmpty(plainText))
                {
                    byte[] binaryRichText = Encoding.Unicode.GetBytes(richText);
                    md5.TransformBlock(binaryRichText, 0, binaryRichText.Length, binaryRichText, 0);
                    byte[] binaryHtml = Encoding.Unicode.GetBytes(htmlText);
                    md5.TransformFinalBlock(binaryHtml, 0, binaryHtml.Length);
                }
                else
                {
                    byte[] binaryType = Encoding.Unicode.GetBytes(typeText);
                    md5.TransformFinalBlock(binaryType, 0, binaryType.Length);
                }
                hash = Convert.ToBase64String(md5.Hash);
                //SQLiteCommand commandSelect = new SQLiteCommand(sql, m_dbConnection);
                //commandSelect.Parameters.AddWithValue("@Hash", hash);

                var oldCurrentClip = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Hash.Equals(hash));

                if (oldCurrentClip != null)
                {
                    oldCurrentClipId = oldCurrentClip.Id;
                    if (true  && lastPastedClips.ContainsKey(oldCurrentClipId) && DateDiffMilliseconds(lastPastedClips[oldCurrentClipId], dtNow) < 1000) // Protection from automated return copy after we send paste. For example Word does so for html paste.
                    {
                        return false;
                    }
                    used = oldCurrentClip.Used;
                    favorite = oldCurrentClip.Favorite;
                    clipTitle = oldCurrentClip.Title;

                    if (true && oldCurrentClipId == LastId && msFromLastCapture > 100) // Protection from automated repeated copy. For example PuntoSwitcher does so.
                    {
                        lastClipWasMultiCaptured = true;
                    }

                    var oldCurrentClipDeleted = liteDb.GetCollection<Clip>("clips").DeleteMany(x => x.Id == oldCurrentClipId); ;
                    RegisterClipIdChange(oldCurrentClipId, LastId + 1);

                }

                //using (SQLiteDataReader reader = commandSelect.ExecuteReader())
                //{
                //    if (reader.Read())
                //    {
                //        oldCurrentClipId = reader.GetInt32(reader.GetOrdinal("Id"));
                //        if (true
                //            && lastPastedClips.ContainsKey(oldCurrentClipId)
                //            && DateDiffMilliseconds(lastPastedClips[oldCurrentClipId], dtNow) < 1000) // Protection from automated return copy after we send paste. For example Word does so for html paste.
                //        {
                //            return false;
                //        }
                //        used = GetNullableBoolFromSqlReader(reader, "Used");
                //        favorite = GetNullableBoolFromSqlReader(reader, "Favorite");
                //        clipTitle = reader.GetString(reader.GetOrdinal("Title"));
                //        sql = "DELETE FROM Clips Where Id = @Id";
                //        SQLiteCommand commandDelete = new SQLiteCommand(sql, m_dbConnection);
                //        if (true
                //            && oldCurrentClipId == LastId
                //            && msFromLastCapture > 100) // Protection from automated repeated copy. For example PuntoSwitcher does so.
                //        {
                //            lastClipWasMultiCaptured = true;
                //        }
                //        commandDelete.Parameters.AddWithValue("@Id", oldCurrentClipId);
                //        commandDelete.ExecuteNonQuery();
                //        RegisterClipIdChange(oldCurrentClipId, LastId + 1);
                //    }
                //}
            }
            else
            {
                Guid g = Guid.NewGuid();
                hash = Convert.ToBase64String(g.ToByteArray());
            }
            LastId++;
            lastClips.Add(new LastClip { Created = created, ID = LastId, ProcessID = 0 });
            int lastClipsMaxSize = 5;
            while (lastClips.Count > lastClipsMaxSize)
            {
                lastClips.RemoveRange(0, lastClips.Count - lastClipsMaxSize);
            }

            //SQLiteCommand commandInsert = new SQLiteCommand("", m_dbConnection);
            //commandInsert.Parameters.AddWithValue("@Id", LastId);
            //commandInsert.Parameters.AddWithValue("@Title", clipTitle);
            //commandInsert.Parameters.AddWithValue("@Text", plainText);
            //commandInsert.Parameters.AddWithValue("@RichText", richText);
            //commandInsert.Parameters.AddWithValue("@HtmlText", htmlText);
            //commandInsert.Parameters.AddWithValue("@Application", applicationText);
            //commandInsert.Parameters.AddWithValue("@Window", windowText);
            //commandInsert.Parameters.AddWithValue("@Created", created);
            //commandInsert.Parameters.AddWithValue("@Type", typeText);
            //commandInsert.Parameters.AddWithValue("@Binary", binaryBuffer);
            //commandInsert.Parameters.AddWithValue("@ImageSample", imageSampleBuffer);
            //commandInsert.Parameters.AddWithValue("@Size", byteSize);
            //commandInsert.Parameters.AddWithValue("@Chars", chars);
            //commandInsert.Parameters.AddWithValue("@Used", used);
            //commandInsert.Parameters.AddWithValue("@Url", url);
            //commandInsert.Parameters.AddWithValue("@Favorite", favorite);
            //commandInsert.Parameters.AddWithValue("@Hash", hash);
            //commandInsert.Parameters.AddWithValue("@appPath", appPath);
            //string intoFieldsText = "", intoParamsText = "";
            //foreach (KeyValuePair<string, string> pair in TextPatterns)
            //{
            //    commandInsert.Parameters.AddWithValue("@Contain_" + pair.Key, Regex.IsMatch(plainText, pair.Value, RegexOptions.IgnoreCase));
            //    intoFieldsText += ", Contain_" + pair.Key;
            //    intoParamsText += ", @Contain_" + pair.Key;
            //}
            //sql = "insert into Clips (Id, Title, Text, Application, Window, Created, Type, Binary, ImageSample, Size, Chars, RichText, HtmlText, Used, Favorite, Url, Hash, appPath"
            //      + intoFieldsText + ") "
            //      + "values (@Id, @Title, @Text, @Application, @Window, @Created, @Type, @Binary, @ImageSample, @Size, @Chars, @RichText, @HtmlText, @Used, @Favorite, @Url, @Hash, @appPath"
            //      + intoParamsText + ")";
            //commandInsert.CommandText = sql;
            //commandInsert.ExecuteNonQuery();


            var col = liteDb.GetCollection<Clip>("clips");
            var clip = new Clip
            {
                Id = LastId,
                Title = clipTitle,
                Type = typeText,
                Text = plainText,
                RichText = richText,
                HtmlText = htmlText,
                Application = applicationText,
                Window = windowText,
                Created = DateTime.Now,
                Binary = binaryBuffer,
                ImageSample = imageSampleBuffer,
                Size = byteSize,
                Chars = chars,
                Used = used,
                Url = url,
                Favorite = favorite,
                Hash = hash,
                AppPath = appPath,
                Contain_email = Regex.IsMatch(plainText, "email", RegexOptions.IgnoreCase),
                Contain_filename = Regex.IsMatch(plainText, "filename", RegexOptions.IgnoreCase),
                Contain_number = Regex.IsMatch(plainText, "number", RegexOptions.IgnoreCase),
                Contain_phone = Regex.IsMatch(plainText, "phone", RegexOptions.IgnoreCase),
                Contain_time = Regex.IsMatch(plainText, "time", RegexOptions.IgnoreCase),
                Contain_url = Regex.IsMatch(plainText, "url", RegexOptions.IgnoreCase),
                Contain_url_image = Regex.IsMatch(plainText, "url_image", RegexOptions.IgnoreCase),
                Contain_url_video = Regex.IsMatch(plainText, "url_video", RegexOptions.IgnoreCase)
            };
            col.Upsert(clip);

            if (updateList)
                ReloadList(false, 0, oldCurrentClipId > 0, null, true);
            if (true
                && applicationText == "ScreenshotReader"
                && IsTextType(typeText)
            )
                ShowForPaste(false, true);

            lastCaptureMoment = DateTime.Now;
            return true;
        }

        private void Delete_Click(object sender = null, EventArgs e = null)
        {
            if (settings.ConfirmationBeforeDelete)
            {
                var confirmResult = MessageBox.Show(this, Properties.Resources.ConfirmDeleteSelectedClips, Properties.Resources.Confirmation, MessageBoxButtons.YesNo);
                if (confirmResult != DialogResult.Yes)
                    return;
            }
            allowRowLoad = false;

            foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
            {
                DataRowView dataRow = (DataRowView)selectedRow.DataBoundItem;
                

                if ((bool)dataRow["Favorite"] == false)
                {
                    int rowId = (int)dataRow["Id"];
                    int col = liteDb.GetCollection<Clip>("clips").DeleteMany(x => x.Id == rowId);


                    if (col == 1)
                        dataGridView.Rows.Remove(selectedRow);
                }
            }

            allowRowLoad = true;
            SelectCurrentRow();
            areDeletedClips = true;
        }

        public static Bitmap ApplicationIcon(string filePath, bool original = true)
        {
            Bitmap originalImage = null;
            Bitmap brightImage = null;
            if (originalIconCache.ContainsKey(filePath))
            {
                originalImage = originalIconCache[filePath];
                brightImage = brightIconCache[filePath];
            }
            else
            {
                if (File.Exists(filePath))
                {
                    Icon smallIcon = null;
                    smallIcon = IconTools.GetIconForFile(filePath, ShellIconSize.SmallIcon);
                    originalImage = smallIcon.ToBitmap();
                    brightImage = (Bitmap)ChangeImageOpacity(originalImage, 0.6f);
                }
                originalIconCache[filePath] = originalImage;
                brightIconCache[filePath] = brightImage;
            }
            if (original)
                return originalImage;
            else
                return brightImage;
        }

        public static Bitmap CopyRectImage(Bitmap bitmap, Rectangle selection)
        {
            //int newBottom = selection.Bottom;
            //if (selection.Bottom > bitmap.Height)
            //    newBottom = bitmap.Height;
            //int newRight = selection.Right;
            //if (selection.Right > bitmap.Width)
            //    newRight = bitmap.Width;
            // TODO check other borders
            // Sometimes Clone() raises strange OutOfMemory exception http://www.codingdefined.com/2015/04/solved-bitmapclone-out-of-memory.html
            //Bitmap RectImage = bitmap.Clone(
            //    new Rectangle(selection.Left, selection.Top, newRight, newBottom), bitmap.PixelFormat);
            Bitmap RectImage = new(selection.Width, selection.Height);
            using (Graphics gph = Graphics.FromImage(RectImage))
            {
                gph.DrawImage(bitmap, new Rectangle(0, 0, RectImage.Width, RectImage.Height), selection, GraphicsUnit.Pixel);
            }
            return RectImage;
        }

        public static Image ChangeImageOpacity(Image originalImage, double opacity)
        {
            // <param name="opacity">Opacity, where 1.0 is no opacity, 0.0 is full transparency</param>
            
            const int bytesPerPixel = 4;
            if ((originalImage.PixelFormat & PixelFormat.Indexed) == PixelFormat.Indexed)
            {
                // Cannot modify an image with indexed colors
                return originalImage;
            }

            Bitmap bmp = (Bitmap)originalImage.Clone();

            // Specify a pixel format.
            PixelFormat pxf = PixelFormat.Format32bppArgb;

            // Lock the bitmap's bits.
            Rectangle rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
            BitmapData bmpData = bmp.LockBits(rect, ImageLockMode.ReadWrite, pxf);

            // Get the address of the first line.
            IntPtr ptr = bmpData.Scan0;

            // Declare an array to hold the bytes of the bitmap.
            // This code is specific to a bitmap with 32 bits per pixels
            // (32 bits = 4 bytes, 3 for RGB and 1 byte for alpha).
            int numBytes = bmp.Width * bmp.Height * bytesPerPixel;
            byte[] argbValues = new byte[numBytes];

            // Copy the ARGB values into the array.
            Marshal.Copy(ptr, argbValues, 0, numBytes);

            // Manipulate the bitmap, such as changing the
            // RGB values for all pixels in the the bitmap.
            for (int counter = 0; counter < argbValues.Length; counter += bytesPerPixel)
            {
                // argbValues is in format BGRA (Blue, Green, Red, Alpha)

                // If 100% transparent, skip pixel
                if (argbValues[counter + bytesPerPixel - 1] == 0)
                    continue;

                int pos = 0;
                pos++; // B value
                pos++; // G value
                pos++; // R value

                argbValues[counter + pos] = (byte)(argbValues[counter + pos] * opacity);
            }

            // Copy the ARGB values back to the bitmap
            Marshal.Copy(argbValues, 0, ptr, numBytes);

            // Unlock the bits.
            bmp.UnlockBits(bmpData);

            return bmp;
        }

        public bool ActivateAndCheckTargetWindow(bool activate = true)
        {
            bool isTargetActive = false;

            if (activate)
            {
                if (!this.TopMost)
                {
                    this.Close();
                }
                else
                {
                    SetForegroundWindow(lastActiveParentWindow);
                }
            }
            int waitStep = 5;
            IntPtr hForegroundWindow = IntPtr.Zero;
            for (int i = 0; i < 200; i += waitStep)
            {
                hForegroundWindow = GetForegroundWindow();
                if (hForegroundWindow != IntPtr.Zero)
                    break;
                Thread.Sleep(waitStep);
            }
            isTargetActive = hForegroundWindow == lastActiveParentWindow;

            return isTargetActive;
        }

        public ClipboardOwner GetClipboardOwnerLockerInfo(bool Locker, bool replaceNullWithLastActive = true)
        {
            IntPtr hwnd;
            ClipboardOwner result = new ClipboardOwner();

            if (Locker)
                hwnd = GetOpenClipboardWindow();
            else
                hwnd = GetClipboardOwner();
            if (hwnd == IntPtr.Zero)
            {
                if (replaceNullWithLastActive)
                    hwnd = lastActiveParentWindow;
                else
                {
                    return result;
                }
            }

            _ = GetWindowThreadProcessId(hwnd, out result.processId);

            Process process1 = Process.GetProcessById(result.processId);
            try
            {
                result.application = process1.ProcessName;
                hwnd = process1.MainWindowHandle;
            }
            catch (Exception)
            {
                return result;
            }
            result.appPath = GetProcessMainModuleFullName(result.processId);
            result.windowTitle = GetWindowTitle(hwnd);

            if (true && hwnd != IntPtr.Zero && (false || String.Compare(result.application, "RDCMan", true) == 0 || String.Compare(result.application, "RDP", true) == 0))
            {
                result.isRemoteDesktop = true;
            }
            return result;
        }

        public string GetSelectedTextOfClips(ref string selectedText, PasteMethod itemPasteMethod = PasteMethod.Null, string DelimiterForTextJoin = null)
        {
            string agregateTextToPaste = "";
            if (itemPasteMethod == PasteMethod.Null)
            {
                selectedText = GetSelectedTextOfClip();
                if (!String.IsNullOrEmpty(selectedText))
                    agregateTextToPaste = selectedText;
            }
            if (String.IsNullOrEmpty(agregateTextToPaste))
            {
                agregateTextToPaste = JoinOrPasteTextOfClips(itemPasteMethod, out _, DelimiterForTextJoin);
            }
            return agregateTextToPaste;
        }

        public void RegisterHotKeys()
        {
            EnumModifierKeys Modifiers;
            Keys Key;
            if (ReadHotkeyFromText(settings.GlobalHotkeyOpenLast, out Modifiers, out Key))
                keyboardHook.RegisterHotKey(Modifiers, Key);
            if (ReadHotkeyFromText(settings.GlobalHotkeyOpenCurrent, out Modifiers, out Key))
                keyboardHook.RegisterHotKey(Modifiers, Key);
            if (ReadHotkeyFromText(settings.GlobalHotkeyOpenFavorites, out Modifiers, out Key))
                keyboardHook.RegisterHotKey(Modifiers, Key);
            if (ReadHotkeyFromText(settings.GlobalHotkeyIncrementalPaste, out Modifiers, out Key))
                keyboardHook.RegisterHotKey(Modifiers, Key);
            if (ReadHotkeyFromText(settings.GlobalHotkeyDecrementalPaste, out Modifiers, out Key))
                keyboardHook.RegisterHotKey(Modifiers, Key);
            if (ReadHotkeyFromText(settings.GlobalHotkeyCompareLastClips, out Modifiers, out Key))
                keyboardHook.RegisterHotKey(Modifiers, Key);
            if (ReadHotkeyFromText(settings.GlobalHotkeyPasteText, out Modifiers, out Key))
                keyboardHook.RegisterHotKey(Modifiers, Key);
            if (ReadHotkeyFromText(settings.GlobalHotkeySimulateInput, out Modifiers, out Key))
                keyboardHook.RegisterHotKey(Modifiers, Key);
            if (ReadHotkeyFromText(settings.GlobalHotkeySwitchMonitoring, out Modifiers, out Key))
                keyboardHook.RegisterHotKey(Modifiers, Key);
            if (ReadHotkeyFromText(settings.GlobalHotkeyForcedCapture, out Modifiers, out Key))
                keyboardHook.RegisterHotKey(Modifiers, Key);
        }

        protected virtual void LoadRowReader(int CurrentRowIndex = -1)
        {
            DataRowView CurrentRowView;
            //LoadedClipRowReader = null;
            if (CurrentRowIndex == -1)
            {
                CurrentRowView = clipBindingSource.Current as DataRowView;
            }
            else
            {
                CurrentRowView = clipBindingSource[CurrentRowIndex] as DataRowView;
            }
            if (CurrentRowView == null)
                return;
            DataRow CurrentRow = CurrentRowView.Row;
            //LoadedClipRowReader = getRowReader((int)CurrentRow["Id"]);
            int rowId = (int)CurrentRow["Id"];


            var clip = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Id == rowId);


            this.clip = clip;

        }

        private static void CalculateByteAndCharSizeOfClip(string htmlText, string richText, string plainText, ref int chars, ref int byteSize)
        {
            if (chars == 0)
                chars = plainText.Length;
            byteSize += plainText.Length * 2; // dirty
            byteSize += htmlText.Length * 2; // dirty
            byteSize += richText.Length * 2; // dirty
        }

        private static string ConvertTextToLine(string agregateTextToPaste)
        {
            agregateTextToPaste = agregateTextToPaste.Trim();
            agregateTextToPaste = Regex.Replace(agregateTextToPaste, "\\s+", " ");
            return agregateTextToPaste;
        }

        private static bool DoActiveWindowBelongsToCurrentProcess(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
                hwnd = GetForegroundWindow();
            _ = GetWindowThreadProcessId(hwnd, out int targetProcessId);
            bool targetIsCurrentProcess = targetProcessId == Environment.ProcessId;
            return targetIsCurrentProcess;
        }

        private static Image GetImageFromBinary(byte[] binary)
        {
            MemoryStream memoryStream = new MemoryStream(binary, 0, binary.Length);
            memoryStream.Write(binary, 0, binary.Length);
            Bitmap image = new(memoryStream);
            //image.MakeTransparent(); // It just makes all black color pixels become transparent
            return image;
        }

        private static string GetWindowTitle(IntPtr hwnd)
        {
            string windowTitle = "";
            //if (settings.ReadWindowTitles)
            //{
            int nChars = Math.Max(1024, GetWindowTextLength(hwnd) + 1);
            StringBuilder buff = new StringBuilder(nChars);
            if (GetWindowText(hwnd, buff, buff.Capacity) > 0)
            {
                windowTitle = buff.ToString();
            }
            //}
            return windowTitle;
        }

        private static bool IsKeyPassedFromFilterToGrid(Keys key, bool isCtrlDown = false)
        {
            return false
                   || key == Keys.Down
                   || key == Keys.Up
                   || key == Keys.PageUp
                   || key == Keys.PageDown
                   || key == Keys.ControlKey
                   || key == Keys.ControlKey
                   || key == Keys.Home && isCtrlDown
                   || key == Keys.End && isCtrlDown;
        }

        private static bool ReadHotkeyFromText(string HotkeyText, out EnumModifierKeys Modifiers, out Keys Key)
        {
            Modifiers = 0;
            Key = 0;
            if (HotkeyText == "" || HotkeyText == "No")
                return false;
            string[] Frags = HotkeyText.Split(new[] { "+" }, StringSplitOptions.None);
            for (int i = 0; i < Frags.Length - 1; i++)
            {
                EnumModifierKeys Modifier;
                Enum.TryParse(Frags[i].Trim(), true, out Modifier);
                Modifiers |= Modifier;
            }
            _ = Enum.TryParse(Frags[Frags.Length - 1], out Key);
            return true;
        }

        private static string RemoveInvalidCharsFromFileName(string fileName, string replacement = "_")
        {
            string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
            fileName = r.Replace(fileName, replacement);
            return fileName;
        }

        private static void SetTextInClipboardDataObject(DataObject dto, string text)
        {
            if (String.IsNullOrEmpty(text))
                return;
            dto.SetText(text, TextDataFormat.Text);
            dto.SetText(text, TextDataFormat.UnicodeText);
        }

        private static string TextClipTitle(string text)
        {
            string title = text.TrimStart();
            // Removing repeats (series) of empty space and leave only 1 space
            title = Regex.Replace(title, @"\s+", " ");
            if (title.Length > ClipTitleLength)
            {
                // Removing repeats (series of one char) of non digits and leave only 8 chars
                title = Regex.Replace(title, "([^\\d])(?<=\\1\\1\\1\\1\\1\\1\\1\\1\\1)", String.Empty,
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
                // Removing repeats (series of one char) of digits and leave only 20 chars
                title = Regex.Replace(title, "(\\d)(?<=\\1\\1\\1\\1\\1\\1\\1\\1\\1\\1\\1\\1\\1\\1\\1\\1\\1\\1\\1\\1\\1)", String.Empty,
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
            }
            if (title.Length > ClipTitleLength)
            {
                title = string.Concat(title.AsSpan(0, ClipTitleLength - 1 - 3), "...");
            }
            return title;
        }

        private static string[] TextToLines(string Text)
        {
            string[] tokens = Regex.Split(Text, @"\r?\n|\r");
            return tokens;
        }

        private void activateListToolStripMenuItem_Click(object sender = null, EventArgs e = null)
        {
            if (dataGridView.Focused)
                FocusClipText();
            else
                dataGridView.Focus();
        }

        private void addSelectedTextInFilterToolStripMenu_Click(object sender, EventArgs e)
        {
            AllowFilterProcessing = false;
            if (!String.IsNullOrWhiteSpace(comboBoxSearchString.Text))
                comboBoxSearchString.Text += " ";
            comboBoxSearchString.Text += GetSelectedTextOfClip();
            AllowFilterProcessing = true;
            SearchStringApply();
        }

        private void AfterRowLoad(bool FullTextLoad = false, int CurrentRowIndex = -1, int NewSelectionStart = -1,  int NewSelectionLength = -1)
        {
            //DataRowView CurrentRowView;
            IHTMLDocument2 htmlDoc;
            string clipType;
            string textPattern = RegexpPattern();
            bool autoSelectMatch = (textPattern.Length > 0 && settings.AutoSelectMatch);
            FullTextLoad = FullTextLoad || EditMode;
            richTextBox.ReadOnly = !EditMode;
            FilterMatches = null;
            bool useNativeTextFormatting = false;
            htmlMode = false;
            bool elementPanelHasFocus = false
                                        || ImageControl.Focused
                                        || richTextBox.Focused
                                        || htmlTextBox.Focused
                                        || urlTextBox.Focused;
            allowTextPositionChangeUpdate = false;
            clipRichTextLength = 0;
            clipType = "";
            pictureBoxSource.Image = null;
            ImageControl.Image = null;
            if (settings.MonospacedFont)
                richTextBox.Font = new Font(FontFamily.GenericMonospace, settings.richTextBoxFontSize);
            else
                richTextBox.Font = new Font(settings.richTextBoxFontFamily, settings.richTextBoxFontSize);
            richTextInternal.Font = richTextBox.Font;
            int fontsize = (int)richTextBox.Font.Size; // Size should be without digits after comma
            richTextBox.SelectionTabs = new int[] { fontsize * 4, fontsize * 8, fontsize * 12, fontsize * 16 }; // Set tab size ~ 4
            richTextInternal.SelectionTabs = richTextBox.SelectionTabs;

            richTextInternal.Clear();
            urlTextBox.Text = "";
            textBoxApplication.Text = "";
            textBoxWindow.Text = "";
            StripLabelCreated.Text = "";
            StripLabelSize.Text = "";
            StripLabelVisualSize.Text = "";
            StripLabelType.Text = "";
            stripLabelPosition.Text = "";

            DataRowView CurrentRowView;
            if (CurrentRowIndex == -1)
            {
                CurrentRowView = clipBindingSource.Current as DataRowView;
            }
            else
            {
                CurrentRowView = clipBindingSource[CurrentRowIndex] as DataRowView;
            }
            if (CurrentRowView == null)
                return;
            DataRow CurrentRow = CurrentRowView.Row;


            int rowId = (int)CurrentRow["Id"];
            var clip = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Id == rowId);

            LoadRowReader();
            if (true && this.clip != null  && !(clip.Created.ToString().Length == 0)) // protection from reading deleted clip  )
            {
            clipType = clip.Type;
            string fullText = clip.Text;
            string fullRTF = clip.RichText;
            string htmlText = GetHtmlFromHtmlClipText();
            //string htmlText = "";
            useNativeTextFormatting = true && settings.ShowNativeTextFormatting && (clipType == "html" || clipType == "rtf");

                Bitmap appIcon = null;
                if (clip.AppPath != null)
                {
                    appIcon = ApplicationIcon(clip.AppPath);
                }
                    

                htmlDoc = htmlTextBox.Document.DomDocument as mshtml.IHTMLDocument2;
                if (clipType == "html")
                {
                    htmlTextBox.Parent.Enabled = false; // Prevent stealing focus
                    htmlDoc.write("");
                    htmlDoc.close(); // Steals focus!!!
                }
                if (appIcon != null)
                {
                    pictureBoxSource.Image = appIcon;
                }
                textBoxApplication.Text = clip.Application;
                textBoxWindow.Text = clip.Window;
                StripLabelCreated.Text = clip.Created.ToString();
                if (!(clip.Size > 0))
                    StripLabelSize.Text = clip.Size.ToString() + " B";
                //StripLabelSize.Text = FormattedClipNumericPropery((int)clip.Size, Properties.Resources.ByteUnit);
                if (!(clip.Chars > 0))
                    StripLabelVisualSize.Text = clip.Chars.ToString() + " C";
                //StripLabelVisualSize.Text = FormattedClipNumericPropery((int)clip.Chars, Properties.Resources.CharUnit);
                StripLabelType.Text = clipType;
                stripLabelPosition.Text = "1";
                string shortText;
                string endMarker;
                Font markerFont = richTextBox.Font;
                Color markerColor;
                if (!FullTextLoad && MaxTextViewSize < fullText.Length)
                {
                    shortText = fullText.Substring(0, MaxTextViewSize);
                    richTextInternal.Text = shortText;
                    endMarker = "<" + Properties.Resources.CutMarker + ">";
                    markerFont = new Font(markerFont, FontStyle.Underline);
                    TextWasCut = true;
                    markerColor = Color.Blue;
                }
                else
                {
                    if (useNativeTextFormatting)
                    {
                        htmlMode = true && clipType == "html" && !string.IsNullOrEmpty(htmlText);
                        if (htmlMode)
                        {
                            string marker = "<DIV class=\"fullSize fdFieldMainContainer";
                            string replacement = "<DIV style=\"width: 100%\" class=\"fullSize fdFieldMainContainer";
                            htmlText = htmlText.Replace(marker, replacement);

                            string newStyle = " margin: 0;";
                            if (settings.WordWrap)
                            {
                                newStyle += " word-wrap: break-word;";
                            }
                            htmlText = Regex.Replace(htmlText, "<body", "<body style=\"" + newStyle + "\" ", RegexOptions.IgnoreCase);
                            htmlDoc.write(htmlText);
                            htmlDoc.close(); // Steals focus!!!
                            mshtml.IHTMLBodyElement body = htmlDoc.body as mshtml.IHTMLBodyElement;
                            htmlTextBox.Document.Body.Drag += new HtmlElementEventHandler(htmlTextBoxDrag); // to prevent internal drag&drop
                            htmlTextBox.Document.Body.KeyDown += new HtmlElementEventHandler(htmlTextBoxDocumentKeyDown);

                            // Need to be called every time, else handler will be lost
                            htmlTextBox.Document.AttachEventHandler("onselectionchange", this.htmlTextBoxDocumentSelectionChange); // No multi call to handler, but why?
                            if (!htmlInitialized)
                            {
                                mshtml.HTMLDocumentEvents2_Event iEvent = (mshtml.HTMLDocumentEvents2_Event)htmlDoc;
                                iEvent.onclick += new mshtml.HTMLDocumentEvents2_onclickEventHandler(htmlTextBoxDocumentClick); //
                                iEvent.onmousedown += new mshtml.HTMLDocumentEvents2_onmousedownEventHandler(htmlTextBoxMouseDown); //
                                htmlInitialized = true;
                            }
                        }
                        else
                        {
                            richTextInternal.Rtf = fullRTF;
                        }
                    }
                    else
                        richTextInternal.Text = fullText;
                    endMarker = "<" + Properties.Resources.EndMarker + ">";
                    TextWasCut = false;
                    markerColor = Color.Green;
                }
                clipRichTextLength = richTextInternal.TextLength;
                if (!EditMode)
                {
                    richTextInternal.SelectionStart = richTextInternal.TextLength;
                    richTextInternal.SelectionColor = markerColor;
                    richTextInternal.SelectionFont = markerFont;
                    if (TextWasCut)
                        endMarker = Environment.NewLine + endMarker;
                    richTextInternal.AppendText(endMarker);
                    // Do it first, else ending hyperlink will connect underline to it

                    MarkLinksInRichTextBox(richTextInternal, out TextLinkMatches);
                    if (textPattern.Length > 0)
                    {
                        MarkRegExpMatchesInRichTextBox(richTextInternal, textPattern, Color.Red, true, false, !string.IsNullOrEmpty(searchString), out FilterMatches);
                    }
                    richTextInternal.AppendText(Environment.NewLine); // adding new line to prevent horizontal scroll to end of extra long last line
                }
                richTextInternal.SelectionColor = new Color();
                richTextInternal.SelectionStart = 0;
                richTextBox.Rtf = richTextInternal.Rtf;

                urlTextBox.HideSelection = true;
                urlTextBox.Clear();
                urlTextBox.Text = clip.Url;
                MarkLinksInRichTextBox(urlTextBox, out UrlLinkMatches);
                if (clipType == "html")
                {
                    htmlTextBox.Parent.Enabled = true;
                }
            }
            else
            {
                richTextBox.Clear();
            }
            tableLayoutPanelData.SuspendLayout();
//            //UpdateClipButtons();
            if (comboBoxSearchString.Focused)
            {
                // Antibug webBrowser steals focus. We set it back
                int filterSelectionLength = comboBoxSearchString.SelectionLength;
                int filterSelectionStart = comboBoxSearchString.SelectionStart;
                comboBoxSearchString.Focus();
                comboBoxSearchString.SelectionStart = filterSelectionStart;
                comboBoxSearchString.SelectionLength = filterSelectionLength;
            }
            if (clipType == "img")
            {
                tableLayoutPanelData.RowStyles[0].Height = 25;
                tableLayoutPanelData.RowStyles[0].SizeType = SizeType.Absolute;
                tableLayoutPanelData.RowStyles[1].Height = 0;
                tableLayoutPanelData.RowStyles[1].SizeType = SizeType.Absolute;
                tableLayoutPanelData.RowStyles[2].Height = 100;
                tableLayoutPanelData.RowStyles[2].SizeType = SizeType.Percent;
                if (elementPanelHasFocus)
                    ImageControl.Focus();
                htmlTextBox.Visible = false; // Without it htmlTextBox will be visible but why?
            }
            else if (htmlMode)
            {
                tableLayoutPanelData.RowStyles[0].Height = 0;
                tableLayoutPanelData.RowStyles[0].SizeType = SizeType.Absolute;
                tableLayoutPanelData.RowStyles[1].Height = 100;
                tableLayoutPanelData.RowStyles[1].SizeType = SizeType.Percent;
                tableLayoutPanelData.RowStyles[2].Height = 0;
                tableLayoutPanelData.RowStyles[2].SizeType = SizeType.Absolute;
                //htmlTextBox.Enabled = true;
                if (elementPanelHasFocus)
                    htmlTextBox.Document.Focus();
                htmlTextBox.Visible = true;
            }
            else
            {
                tableLayoutPanelData.RowStyles[0].Height = 100;
                tableLayoutPanelData.RowStyles[0].SizeType = SizeType.Percent;
                tableLayoutPanelData.RowStyles[1].Height = 0;
                tableLayoutPanelData.RowStyles[1].SizeType = SizeType.Percent;
                tableLayoutPanelData.RowStyles[2].Height = 0;
                tableLayoutPanelData.RowStyles[2].SizeType = SizeType.Absolute;
                if (elementPanelHasFocus)
                    richTextBox.Focus();
            }
            if (urlTextBox.Text == "")
                tableLayoutPanelData.RowStyles[3].Height = 0;
            else
                tableLayoutPanelData.RowStyles[3].Height = 25;
            if (EditMode && this.Visible && this.ContainsFocus)
                richTextBox.Focus(); // Can activate this window, so we check that window has focus
            if (clipType == "img")
            {
                Image image = GetImageFromBinary((byte[])clip.Binary);
                ImageControl.Image = image;
            }
            tableLayoutPanelData.ResumeLayout();
            if (clipType == "img")
            {
                ImageControl.ZoomFitInside();
            }
            else
            {
                if (autoSelectMatch)
                    SelectNextMatchInClipText(false);
                else
                {
                    if (NewSelectionStart == -1)
                        NewSelectionStart = selectionStart;
                    if (NewSelectionLength == -1)
                        NewSelectionLength = selectionLength;
                    if (NewSelectionStart > 0 || NewSelectionLength > 0)
                    {
                        if (htmlMode)
                        {
                            IHTMLTxtRange range = SelectTextRangeInWebBrowser(NewSelectionStart, NewSelectionLength);
                            range.scrollIntoView();
                        }
                        else
                        {
                            SetRichTextboxSelection(NewSelectionStart, NewSelectionLength, true);
                        }
                    }
                }
                allowTextPositionChangeUpdate = true;
                OnClipContentSelectionChange();
            }
            tableLayoutPanelData.Refresh(); // to let user see clip in autoSelectNextClip state
        }

        private bool BoolFieldValue(string fieldName, DataRowView dataRowView = null)
        {
            // for nullable bool fields
            if (dataRowView == null)
            {
                dataRowView = (DataRowView)(clipBindingSource.Current);
                //if (dataRowView != null && clip != null)
                //    dataRowView[fieldName] = clip[fieldName]; // DataBoundItem can be not read yet
            }
            bool favVal = false;
            if (dataRowView != null)
            {
                var favVal1 = dataRowView.Row[fieldName];
                if (favVal1 != null)
                    favVal = (bool)favVal1;
            }
            return favVal;
        }

        private void buttonFindNext_Click(object sender = null, EventArgs e = null)
        {
            if (EditMode)
                return;
            SaveFilterInLastUsedList();
            if (TextWasCut)
                AfterRowLoad(true);
            SelectNextMatchInClipText();
        }

        private void buttonFindPrevious_Click(object sender, EventArgs e)
        {
            if (EditMode)
                return;
            SaveFilterInLastUsedList();
            if (TextWasCut)
                AfterRowLoad(true);
            if (htmlMode)
                SelectNextMatchInWebBrowser(-1);
            else
            {
                RichTextBox control = richTextBox;
                if (FilterMatches == null)
                    return;
                Match prevMatch = null;
                foreach (Match match in FilterMatches)
                {
                    if (false
                        || control.SelectionStart > match.Groups[1].Index
                        || (true
                            && control.SelectionLength == 0
                            && match.Index == 0
                        ))
                    {
                        prevMatch = match;
                    }
                }
                if (prevMatch != null)
                {
                    control.SelectionStart = prevMatch.Groups[1].Index;
                    control.SelectionLength = prevMatch.Groups[1].Length;
                    control.HideSelection = false;
                }
            }
        }

        private void CaptureClipboardData()
        {
            DateTime now = DateTime.Now;
            captureTimer.Stop();
            if (tempCaptureTimer.Enabled)
            {
                tempCaptureTimer.Stop();
                settings.MonitoringClipboard = false;
                RemoveClipboardFormatListener(this.Handle);
            }

            IDataObject iData = new DataObject();
            string clipType = "";
            string clipText = "";
            string richText = "";
            string htmlText = "";
            string clipUrl = "";
            string imageUrl = "";
            int clipChars = 0;
            bool needUpdateList = false;
            ClipboardOwner clipboardOwner = GetClipboardOwnerLockerInfo(false);
            if (ignoreModulesInClipCapture.Contains(clipboardOwner.application.ToLower()))
                return;
            try
            {
                iData = Clipboard.GetDataObject();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(String.Format("Clipboard.GetDataObject(): InteropServices.ExternalException: {0}", ex.Message));
                string Message = Properties.Resources.ErrorReadingClipboard + ": " + ex.Message;
                notifyIcon.ShowBalloonTip(2000, Application.ProductName, Message, ToolTipIcon.Info);
                return;
            }
            if (iData.GetDataPresent(DataFormat_RemoveTempClipsFromHistory))
            {
                removeClipsFilter removeClipsFilter;
                string dataString = (string)iData.GetData(DataFormat_RemoveTempClipsFromHistory);
                if (!String.IsNullOrWhiteSpace(dataString))  // Rarely we got NULL
                {
                    removeClipsFilter = System.Text.Json.JsonSerializer.Deserialize<removeClipsFilter>(dataString);
                    double maxRemovedClipAge = removeClipsFilter.TimeEnd.Subtract(removeClipsFilter.TimeStart).TotalMilliseconds; // Докинем немного времени на путь от источника до приемника
                    double timeDelta = now.Subtract(removeClipsFilter.TimeEnd).TotalMilliseconds;
                    for (int i = lastClips.Count - 1; i >= 0; i--)
                    {
                        LastClip lastClip = lastClips[i];
                        double clipAge = now.Subtract(lastClip.Created).TotalMilliseconds;
                        if (true && clipAge < maxRemovedClipAge)
                        {
                            liteDb.GetCollection<Clip>("clips").DeleteMany(x => x.Id== lastClip.ID);
                            lastClips.RemoveAt(i);
                            needUpdateList = true;
                        }
                    }
                }
            }
            if (iData.GetDataPresent(DataFormat_ClipboardViewerIgnore) && settings.IgnoreExclusiveFormatClipCapture)
            {
                if (needUpdateList)
                    ReloadList();
                return;
            }
            bool textFormatPresent = false;
            byte[] binaryBuffer = Array.Empty<byte>();
            byte[] imageSampleBuffer = Array.Empty<byte>();
            int NumberOfFilledCells = 0;
            int NumberOfImageCells = 0;
            Bitmap bitmap = null;
            try
            {
                if (iData.GetDataPresent(DataFormats.UnicodeText))
                {
                    clipText = (string)iData.GetData(DataFormats.UnicodeText);
                    if (!String.IsNullOrEmpty(clipText))
                    {
                        clipType = "text";
                        textFormatPresent = true;
                    }
                }
                if (!textFormatPresent && iData.GetDataPresent(DataFormats.Text))
                {
                    clipText = (string)iData.GetData(DataFormats.Text);
                    if (!String.IsNullOrEmpty(clipText))
                    {
                        clipType = "text";
                        textFormatPresent = true;
                    }
                }

                if (iData.GetDataPresent(DataFormat_XMLSpreadSheet))
                {
                    object data = iData.GetData(DataFormat_XMLSpreadSheet);
                    if (data.GetType() == typeof(MemoryStream))
                    {
                        // Excel
                        using MemoryStream ms = (MemoryStream)data;
                        if (ms != null && ms.Length > 0)
                        {
                            byte[] buffer = new byte[ms.Length];
                            ms.Read(buffer, 0, (int)ms.Length);
                            string xmlSheet = Encoding.Default.GetString(buffer);
                            Match match = Regex.Match(xmlSheet, "ExpandedColumnCount=\"(\\d+)\"");
                            int NumberOfColumns = 1;
                            if (match.Success)
                                NumberOfColumns = Convert.ToInt32(match.Groups[1].Value);
                            match = Regex.Match(xmlSheet, "ExpandedRowCount=\"(\\d+)\"");
                            int NumberOfRows = 1;
                            if (match.Success)
                                NumberOfRows = Convert.ToInt32(match.Groups[1].Value);
                            NumberOfImageCells = NumberOfRows * NumberOfColumns;
                            NumberOfFilledCells = Regex.Matches(xmlSheet, "<Row").Count;
                        }
                    }
                }
                if (true && iData.GetDataPresent(DataFormats.Html) && (false || NumberOfFilledCells == 0 || settings.MaxCellsToCaptureFormattedText > NumberOfFilledCells))
                {
                    htmlText = (string)iData.GetData(DataFormats.Html);
                    if (String.IsNullOrEmpty(htmlText))
                    {
                        htmlText = "";
                    }
                    else
                    {
                        clipType = "html";

                        Match match = Regex.Match(htmlText, @"SourceURL:(file:///?)?(.*?)(?:\n|\r|$)", RegexOptions.IgnoreCase);
                        if (match.Captures.Count > 0)
                        {
                            clipUrl = match.Groups[2].ToString();
                            if (!String.IsNullOrEmpty(match.Groups[1].ToString()))
                            {
                                clipUrl = System.Web.HttpUtility.UrlDecode(clipUrl);
                                clipUrl = clipUrl.Replace(@"/", @"\");
                            }
                        }
                    }
                }

                if (true && iData.GetDataPresent(DataFormats.Rtf) && clipType != "html"  && (false || NumberOfFilledCells == 0 || settings.MaxCellsToCaptureFormattedText > NumberOfFilledCells))
                {
                    richText = (string)iData.GetData(DataFormats.Rtf);
                    clipType = "rtf";
                    if (!textFormatPresent)
                    {
                        var rtfBox = new RichTextBox();
                        rtfBox.Rtf = richText;
                        clipText = rtfBox.Text;
                        textFormatPresent = true;
                    }
                }
                if (settings.CaptureImages && textFormatPresent && bitmap == null)
                {
                    Match match;
                    match = Regex.Match(clipText, "^\\s*" + videoPattern + "\\s*$", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        imageUrl = "http://img.youtube.com/vi/";
                        int youtubeIndex = (match.Groups.Count - 1);
                        var youtubeId = match.Groups[youtubeIndex].ToString();
                        imageUrl = imageUrl + youtubeId + "/default.jpg";
                        bitmap = getBitmapFromUrl(imageUrl);
                    }
                    match = Regex.Match(clipText, "^\\s*" + imagePattern + "\\s*$", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        imageUrl = match.Value;
                        bitmap = getBitmapFromUrl(imageUrl);
                    }
                }

                if (iData.GetDataPresent(DataFormats.FileDrop))
                {
                    if (iData.GetData(DataFormats.FileDrop) is string[] fileNameList)
                    {
                        if (settings.CaptureImages && fileNameList.Length == 1 && iData.GetDataPresent(DataFormats.Bitmap))
                        {
                            // Command "Copy image" executed in browser IE
                            clipType = "";
                        }
                        else
                        {
                            clipText = String.Join("\n", fileNameList);
                            clipType = "file";
                        }
                    }
                    else
                    {
                        // Coping Outlook task
                    }
                }

                string clipTextImage = "";
                int clipCharsImage = 0;
                // http://www.cyberforum.ru/ado-net/thread832314.html
                // html text check to prevent crush from too big generated Excel image
                if (true && settings.CaptureImages && iData.GetDataPresent(DataFormats.Bitmap)
                    && (false
                        || NumberOfImageCells == 0 && string.IsNullOrWhiteSpace(clipText)
                        || NumberOfImageCells != 0 && settings.MaxCellsToCaptureImage > NumberOfImageCells))
                {
                    //clipType = "img";
                    
                    if (iData.GetDataPresent(DataFormats.Dib))
                    {
                        // First - try get 24b bitmap + 8b alpfa (transparency)
                        // https://www.csharpcodi.com/vs2/1561/noterium/src/Noterium.Core/Helpers/ClipboardHelper.cs/
                        // https://www.hostedredmine.com/issues/929403
                        var dibData = Clipboard.GetData(DataFormats.Dib);
                        if (dibData != null)
                        {
                            var dib = ((MemoryStream)dibData).ToArray();
                            var width = BitConverter.ToInt32(dib, 4);
                            var height = BitConverter.ToInt32(dib, 8);
                            var bpp = BitConverter.ToInt16(dib, 14);
                            if (bpp == 32)
                            {
                                var gch = GCHandle.Alloc(dib, GCHandleType.Pinned);
                                try
                                {
                                    var ptr = new IntPtr((long)gch.AddrOfPinnedObject() + 40);
                                    bitmap = new Bitmap(width, height, width * 4, PixelFormat.Format32bppArgb, ptr);
                                    bitmap.RotateFlip(RotateFlipType.Rotate180FlipX);
                                }
                                finally
                                {
                                    gch.Free();
                                }
                            }
                        }
                    }

                    if (bitmap == null)
                        // Second - get 24b bitmap
                        bitmap = iData.GetData(DataFormats.Bitmap, false) as Bitmap;


                    if (bitmap != null) // NUll happens while copying image in standard image viewer Windows 10
                    {
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            bitmap.Save(memoryStream, ImageFormat.Png);
                            binaryBuffer = memoryStream.ToArray();
                        }
                        //clipTextImage = Properties.Resources.Size + ": " + image.Width + "x" + image.Height + "\n"
                        //     + Properties.Resources.PixelFormat + ": " + image.PixelFormat + "\n";
                        clipTextImage = bitmap.Width + " x " + bitmap.Height;
                        if (!String.IsNullOrEmpty(clipboardOwner.windowTitle))
                            clipTextImage += ", " + clipboardOwner.windowTitle;
                        clipTextImage += ", " + Properties.Resources.PixelFormat + ": " + Image.GetPixelFormatSize(bitmap.PixelFormat);
                        clipCharsImage = bitmap.Width * bitmap.Height;
                    }
                }
                if (bitmap != null)
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        int fragmentWidth = 100;
                        int fragmentHeight = 20;
                        var bestPoint = FindBestImageFragment(bitmap, fragmentWidth, fragmentHeight);
                        using (Image ImageSample = CopyRectImage(bitmap, new Rectangle(bestPoint.X, bestPoint.Y, fragmentWidth, fragmentHeight)))
                        {
                            ImageSample.Save(memoryStream, ImageFormat.Png);
                            imageSampleBuffer = memoryStream.ToArray();
                        }
                    }

                if (clipType == "html" && clipText == "")
                    clipType = "";
                // Split Image+Html clip into 2: Image and Html
                if (clipTextImage != "")
                {
                    // Image clip
                    if (!String.IsNullOrEmpty(clipUrl))
                        imageUrl = clipUrl;
                    if (imageUrl.StartsWith("data:image"))
                        imageUrl = "";
                    bool clipAdded = AddClip(binaryBuffer, imageSampleBuffer, "", "", "img", clipTextImage, clipboardOwner.application, clipboardOwner.windowTitle, imageUrl, clipCharsImage, clipboardOwner.appPath, false, false, false, "");
                    needUpdateList = needUpdateList || clipAdded;

                    if (!String.IsNullOrWhiteSpace(clipText))
                        imageSampleBuffer = Array.Empty<byte>();
                }
                if (clipType != "")
                {
                    // Non image clip
                    bool clipAdded = AddClip(Array.Empty<byte>(), imageSampleBuffer, htmlText, richText, clipType, clipText, clipboardOwner.application, clipboardOwner.windowTitle, clipUrl, clipChars, clipboardOwner.appPath, false, false, false, "");
                    needUpdateList = needUpdateList || clipAdded;
                }
            }
            finally
            {
                bitmap?.Dispose();
            }
            if (needUpdateList)
                ReloadList();
        }

        private void ClearFilter(int CurrentClipID = 0, bool keepMarkFilterFilter = false, bool waitFinish = false)
        {
            if (filterOn)
            {
                AllowFilterProcessing = false;
                comboBoxSearchString.Text = "";
                periodFilterOn = false;
                ReadSearchString();
                TypeFilter.SelectedIndex = 0;
                if (!keepMarkFilterFilter)
                    MarkFilter.SelectedIndex = 0;
                AllowFilterProcessing = true;
                //UpdateClipBindingSource(false, CurrentClipID);
                ReloadList(true, CurrentClipID, false, null, waitFinish); // To repaint text
            }
            else if (CurrentClipID != 0)
                RestoreSelectedCurrentClip(false, CurrentClipID);
        }

        private void ClearFilter_Click(object sender = null, EventArgs e = null)
        {
            ClearFilter();
            dataGridView.Focus();
        }

        private DataObject ClipDataObject(Clip localClip, bool onlySelectedPlainText, out string clipText)
        {
            clipText = "";
            if (localClip == null)
                localClip = clip;
            if (localClip == null)
                return null;

            //DataRow CurrentDataRow = ((DataRowView)clipBindingSource.Current).Row;
            string type = localClip.Type;
            object richText = localClip.RichText;
            object htmlText = localClip.HtmlText;
            byte[] binary = localClip.Binary;
            clipText = localClip.Text;
            if (type == "img" && onlySelectedPlainText)
            {
                clipText = GetClipTempFile(out _, localClip);
            }
            if (localClip == clip)
            {
                string selectedText = GetSelectedTextOfClip(onlySelectedPlainText);
                if (!String.IsNullOrEmpty(selectedText))
                    clipText = selectedText;
            }
            DataObject dto = new DataObject();
            if (IsTextType(type) || type == "file" || onlySelectedPlainText)
            {
                SetTextInClipboardDataObject(dto, clipText);
            }
            if (true && (type == "rtf" || type == "html") && !(richText == null) && !String.IsNullOrEmpty(richText.ToString()) && !onlySelectedPlainText)
            {
                dto.SetText((string)richText, TextDataFormat.Rtf);
            }
            if (true && type == "html" && htmlText is not DBNull && !onlySelectedPlainText)
            {
                dto.SetText((string)htmlText, TextDataFormat.Html);
            }
            if (type == "file" && !onlySelectedPlainText)
            {
                string[] fileNameList = clipText.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                StringCollection fileNameCollection = new();
                foreach (string fileName in fileNameList)
                {
                    fileNameCollection.Add(fileName);
                }
                dto.SetFileDropList(fileNameCollection);
            }
            if (type == "img" && !onlySelectedPlainText)
            {
                Image image = GetImageFromBinary(binary);
                dto.SetImage(image);
            }
            return dto;
        }

        private string clipTempFile(Clip rowReader, string suffix = "")
        {
            string extension;
            string type = rowReader.Type;
            if (type == "text" || type == "file")
                extension = "txt";
            else if (type == "rtf" || type == "html")
                extension = type;
            else if (type == "img")
                extension = "png";
            else
                extension = "dat";
            string tempFolder = settings.ClipTempFileFolder;
            if (!Directory.Exists(tempFolder))
                tempFolder = Path.GetTempPath();
            if (!tempFolder.EndsWith("\\"))
                tempFolder += "\\";
            string tempFile = tempFolder + "Clip_" + rowReader.Id;
            if (!String.IsNullOrEmpty(suffix))
                tempFile += "_" + suffix;
            tempFile += "." + extension;
            try
            {
                using (new StreamWriter(tempFile))
                {
                }
            }
            catch
            {
                tempFile = "";
            }
            return tempFile;
        }

        private int ColorDifference(Color Color1, Color Color2)
        {
            int result = 0;
            result += 30 * Math.Abs(Color1.R - Color2.R);
            result += 59 * Math.Abs(Color1.G - Color2.G);
            result += 11 * Math.Abs(Color1.B - Color2.B);
            result = result / 100;
            return result;
        }

        private string comparatorExeFileName()
        {
            // TODO read paths from registry and let use custom application
            string path;
            path = settings.TextCompareApplication;
            if (!String.IsNullOrEmpty(path) && File.Exists(path))
            {
                return path;
            }

            path = "C:\\Program Files (x86)\\Beyond Compare 3\\BCompare.exe";
            if (File.Exists(path))
            {
                return path;
            }

            path = "C:\\Program Files\\Beyond Compare 3\\BCompare.exe";
            if (File.Exists(path))
            {
                return path;
            }

            path = "C:\\Program Files (x86)\\ExamDiff Pro\\ExamDiff.exe";
            if (File.Exists(path))
            {
                return path;
            }

            path = "C:\\Program Files\\ExamDiff Pro\\ExamDiff.exe";
            if (File.Exists(path))
            {
                return path;
            }

            path = "C:\\Program Files (x86)\\WinMerge\\WinMergeU.exe";
            if (File.Exists(path))
            {
                return path;
            }

            path = "C:\\Program Files (x86)\\Araxis\\Araxis Merge\\compare.exe";
            if (File.Exists(path))
            {
                return path;
            }

            path = "C:\\Program Files\\Araxis\\Araxis Merge\\compare.exe";
            if (File.Exists(path))
            {
                return path;
            }

            path = "C:\\Program Files (x86)\\SourceGear\\Common\\DiffMerge\\sgdm.exe";
            if (File.Exists(path))
            {
                return path;
            }

            path = "C:\\Program Files\\SourceGear\\Common\\DiffMerge\\sgdm.exe";
            if (File.Exists(path))
            {
                return path;
            }

            path = "C:\\Program Files (x86)\\KDiff3\\kdiff3.exe";
            if (File.Exists(path))
            {
                return path;
            }

            path = "C:\\Program Files\\KDiff3\\KDiff3.exe";
            if (File.Exists(path))
            {
                return path;
            }

            MessageBox.Show(this, Properties.Resources.NoSupportedTextCompareApplication, Application.ProductName);
            Process.Start("http://winmerge.org/");
            return "";
        }

        private void CompareClipsByID(int id1, int id2)
        {
            string comparatorName = comparatorExeFileName();
            if (comparatorName == "")
            {
                return;
            }

            string type1;
            string type2;

            var rowReader1 = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Id==id1);
            var rowReader2 = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Id==id2);

            type1 = rowReader1.Type;
            type2 = rowReader2.Type;
            if (!IsTextType(type1) && type1 != "file")
                return;
            if (!IsTextType(type2) && type2 != "file")
                return;
            string filename1 = clipTempFile(rowReader1, "comp");
            File.WriteAllText(filename1, rowReader1.Text, Encoding.UTF8);
            string filename2 = clipTempFile(rowReader2, "comp");
            File.WriteAllText(filename2, rowReader2.Text, Encoding.UTF8);
            Process.Start(comparatorName, String.Format("\"{0}\" \"{1}\"", filename1, filename2));
        }

        private void ConnectClipboard()
        {
            if (!settings.MonitoringClipboard)
                return;
            if (!AddClipboardFormatListener(this.Handle))
            {
                int ErrorCode = Marshal.GetLastWin32Error();
                int ERROR_INVALID_PARAMETER = 87;
                if (ErrorCode != ERROR_INVALID_PARAMETER)
                    Debug.WriteLine("Failed to connect clipboard: " + Marshal.GetLastWin32Error());
                else
                {
                    //already connected
                }
            }
        }

        private void contextMenuUrlOpenLink_Click(object sender, EventArgs e)
        {
            OpenLinkIfAltPressed(urlTextBox, e, UrlLinkMatches, false);
        }

        private string CopyClipToClipboard(Clip rowReader = null, bool onlySelectedPlainText = false, bool allowSelfCapture = true)
        {
            SaveFilterInLastUsedList();
            string clipText;
            var dto = ClipDataObject(rowReader, onlySelectedPlainText, out clipText);
            if (dto == null)
                return "";
            //if (!settings.MoveCopiedClipToTop)
            //    CaptureClipboard = false;
            SetClipboardDataObject(dto, allowSelfCapture);
            return clipText;
        }

        private bool CurrentIDChanged()
        {
            return false
                   || (clip == null && dataGridView.CurrentRow != null)
                   || (clip != null && dataGridView.CurrentRow == null)
                   || !(true
                        && clip != null
                        && dataGridView.CurrentRow != null
                        && dataGridView.CurrentRow.DataBoundItem != null
                        && (int)(dataGridView.CurrentRow.DataBoundItem as DataRowView)["ID"] == (int)clip.Id);
        }

        private void dataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            var row = dataGridView.Rows[e.RowIndex];
            if (e.ColumnIndex == dataGridView.Columns["VisualWeight"].Index)
            {
                DataGridViewCell hoverCell = row.Cells[e.ColumnIndex];
                if (hoverCell.Value != null)
                    hoverCell.ToolTipText = Properties.Resources.VisualWeightTooltip;
            }
        }

        private void dataGridView_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (!dataGridView.Rows[e.RowIndex].Selected)
                {
                    DataRowView row1 = (DataRowView)dataGridView.Rows[e.RowIndex].DataBoundItem;
                    int newPosition = clipBindingSource.Find("Id", (int)row1["id"]);
                    clipBindingSource.Position = newPosition;
                    SelectCurrentRow();
                }
            }
        }

        private void dataGridView_DoubleClick(object sender, EventArgs e)
        {
            SendPasteOfSelectedTextOrSelectedClips();
        }

        private void dataGridView_KeyDown(object sender, KeyEventArgs e)
        {
            if (true
                && !IsKeyPassedFromFilterToGrid(e.KeyCode, e.Control)
                && e.KeyCode != Keys.Delete
                && e.KeyCode != Keys.Home
                && e.KeyCode != Keys.End
                && e.KeyCode != Keys.Enter
                && e.KeyCode != Keys.ShiftKey
                && e.KeyCode != Keys.Alt
                && e.KeyCode != Keys.Menu
                && e.KeyCode != Keys.Tab
                && e.KeyCode != Keys.Apps
                && e.KeyCode != Keys.F10
                && !e.Control
            )
            {
                comboBoxSearchString.Focus();
                sendKey(comboBoxSearchString.Handle, e.KeyData, false, true);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.A && e.Control)
            {
                // Prevent CTRL+A from selection all clips
                e.Handled = true;
                for (int i = 0; i < Math.Min(dataGridView.RowCount, maxClipsToSelect); i++)
                {
                    dataGridView.Rows[i].Selected = true;
                }
            }
            else if (e.KeyCode == Keys.Insert && e.Control)
            {
                e.Handled = true;
                copyClipToolStripMenuItem_Click();
            }
            else if (e.KeyCode == Keys.Delete)
            {
                e.Handled = true;
                Delete_Click();
            }
        }

        private void dataGridView_MouseMove(object sender, MouseEventArgs e)
        {
            if (LastMousePoint != e.Location)
            {
                LastMousePoint = e.Location;
                if (e.Button == MouseButtons.Left)
                {
                    DataGridView.HitTestInfo info = dataGridView.HitTest(e.X, e.Y);
                    if (info.RowIndex >= 0)
                    {
                        if (info.RowIndex >= 0 && info.ColumnIndex >= 0)
                        {
                            DataObject dto = new DataObject();
                            string agregateTextToPaste = JoinOrPasteTextOfClips(PasteMethod.Null, out _);
                            SetTextInClipboardDataObject(dto, agregateTextToPaste);
                            SetClipFilesInDataObject(dto);
                            dataGridView.DoDragDrop(dto, DragDropEffects.Copy);
                        }
                    }
                }
            }
        }

        private void dataGridView_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            var row = dataGridView.Rows[e.RowIndex];
            if (row.Cells["ColumnTitle"].Value == null)
            {
                PrepareRow(row);
                e.PaintCells(e.ClipBounds, DataGridViewPaintParts.All);
                e.Handled = true;
            }
        }

        private void dataGridView_SelectionChanged(object sender, EventArgs e)
        {
            if(dataGridView.CurrentRow != null && dataGridView.CurrentRow.Cells[10].Value !=null)
            {
                var test = (bool)dataGridView.CurrentRow.Cells[10].Value;
                //toolStripButtonMarkFavorite.Checked = test;
  
                if (test)
                {
                    toolStripMenuItemMarkFavorite.BackColor = Color.LightBlue;
                }
                else
                {
                    toolStripMenuItemMarkFavorite.BackColor = Control.DefaultBackColor;
                }

            }

            if (!allowProcessDataGridSelectionChanged)
                return;
            if (allowRowLoad)
            {
                if (EditMode)
                    toolStripMenuItemEditClipText1_Click();
                else
                    LoadClipIfChangedID();
                SaveFilterInLastUsedList();
            }
            if (dataGridView.Focused || comboBoxSearchString.Focused)
            {
                allowProcessDataGridSelectionChanged = false;
                dataGridView.SuspendLayout();
                if (true && selectedRangeStart >= 0 && dataGridView.CurrentRow != null && (ModifierKeys & Keys.Shift) != 0)
                {
                    
                    // Make natural (2 directions) order of range selected rows
                    int lastIndex = dataGridView.CurrentRow.Index;
                    int firstIndex = Math.Min(selectedRangeStart, dataGridView.RowCount - 1);
                    int step;
                    if (firstIndex > lastIndex)
                        step = -1;
                    else
                        step = +1;
                    for (int i = firstIndex; i != lastIndex + step; i += step)
                    {
                        dataGridView.Rows[i].Selected = false;
                        dataGridView.Rows[i].Selected = true;
                    }
                }
                dataGridView.ResumeLayout();
                if ((ModifierKeys & Keys.Shift) == 0)
                {
                    if (dataGridView.CurrentRow == null)
                        selectedRangeStart = -1;
                    else
                        selectedRangeStart = dataGridView.CurrentRow.Index;
                }
                allowProcessDataGridSelectionChanged = true;


            }
        }

        private int DateDiffMilliseconds(DateTime dt1, object dt2 = null)
        {
            if (dt2 == null)
                dt2 = DateTime.Now;
            TimeSpan span = (DateTime)dt2 - dt1;
            int ms = (int)span.TotalMilliseconds;
            return ms;
        }

        private void deleteAllNonFavoriteClips()
        {
            allowRowLoad = false;

            //string sql = "Delete from Clips where NOT Favorite OR Favorite IS NULL";
            //SQLiteCommand command = new SQLiteCommand("", m_dbConnection);
            //command.CommandText = sql;
            //command.ExecuteNonQuery();

            var col = liteDb.GetCollection<Clip>("clips").DeleteMany(x => x.Favorite != true);
            areDeletedClips = true;
        }

        private void DeleteExcessClips()
        {
            if (settings.HistoryDepthNumber == 0)
                return;
            int clipsCount = liteDb.GetCollection<Clip>("clips").Count();
            int numberOfClipsToDelete = clipsCount - settings.HistoryDepthNumber;
            if (numberOfClipsToDelete > 0)
            {
                var col = liteDb.GetCollection<Clip>("clips").Query()
                    .Where(x => x.Favorite != true)
                    .OrderBy(x => x.Id)
                    .Limit(numberOfClipsToDelete)
                    .ToList();

                foreach (var item in col)
                {
                    var del = liteDb.GetCollection<Clip>("clips").DeleteMany(x => x.Id == item.Id);
                }

            }
        }

        private void DeleteOldClips()
        {
            if (settings.HistoryDepthDays == 0)
                return;

            var col = liteDb.GetCollection<Clip>("clips").Query()
                        .Where(x => x.Favorite != true)
                        .Where(x => x.Created < DateTime.Now.AddDays(- settings.HistoryDepthDays))
                        .OrderBy(x => x.Id)
                        .ToList();

            foreach (var item in col)
            {
                var del = liteDb.GetCollection<Clip>("clips").DeleteMany(x => x.Id == item.Id);
            }

            //SQLiteCommand command = new SQLiteCommand(m_dbConnection);
            //command.CommandText = "Delete From Clips where (NOT Favorite OR Favorite IS NULL) AND Created < date('now','-" + settings.HistoryDepthDays + " day')";
            //commandInsert.Parameters.AddWithValue("Number", settings.HistoryDepthDays);
            //command.Parameters.AddWithValue("CurDate", DateTime.Now);
            //command.ExecuteNonQuery();
        }

        private void exportMenuItem_Click(object sender, EventArgs e)
        {
            //SaveFileDialog saveFileDialog = new SaveFileDialog();
            //saveFileDialog.CheckFileExists = false;
            //saveFileDialog.Filter = "Clippy clips|*.cac|All|*.*";
            //if (saveFileDialog.ShowDialog(this) != DialogResult.OK)
            //    return;
            //string sql = "Select * from Clips where Id IN(null";
            //int counter = 0;
            //foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
            //{
            //    DataRowView dataRow = (DataRowView)selectedRow.DataBoundItem;
            //    string parameterName = "@Id" + counter;
            //    sql += "," + parameterName;
            //    globalDataAdapter.SelectCommand.Parameters.Add(parameterName, DbType.Int32).Value = dataRow["Id"];
            //    counter++;
            //    if (counter == 999) // SQLITE_MAX_VARIABLE_NUMBER, which defaults to 999, but can be lowered at runtime
            //        break;
            //}
            //sql += ")";
            //globalDataAdapter.SelectCommand.CommandText = sql;
            //DataTable dataTable = new DataTable();
            //dataTable.TableName = "ClippyClips";
            //dataTable.Locale = CultureInfo.InvariantCulture;
            //globalDataAdapter.Fill(dataTable);
            //dataTable.WriteXml(saveFileDialog.FileName, XmlWriteMode.WriteSchema);
        }

        private void FillFilterItems()
        {
            int filterSelectionLength = comboBoxSearchString.SelectionLength;
            int filterSelectionStart = comboBoxSearchString.SelectionStart;

            List<string> lastFilterValues = settings.LastFilterValues;
            comboBoxSearchString.Items.Clear();
            foreach (string String in lastFilterValues)
            {
                comboBoxSearchString.Items.Add(String);
            }

            // For some reason selection is reset. So we restore it
            comboBoxSearchString.SelectionStart = filterSelectionStart;
            comboBoxSearchString.SelectionLength = filterSelectionLength;
        }

        private Point FindBestImageFragment(Bitmap bitmap, int fragmentWidth, int fragmentHeight, int diffTreshold = 20)
        {
            int bestX = -1, bestY = -1;
            int goodX = -1, goodY = -1;
            //int badX = -1, badY = -1;
            //int maxDelta = 0;
            for (int x = 0; x < Math.Min(bitmap.Width, 500) - 2; x += 2) // step 1->2 for speed up
            {
                for (int y = 0; y < Math.Min(bitmap.Height, 500) - 2; y += 2) // step 1->2 for speed up
                {
                    Color basePixel = bitmap.GetPixel(x, y);
                    if (true
                        && ColorDifference(basePixel, bitmap.GetPixel(x + 2, y + 1)) > diffTreshold
                        && ColorDifference(basePixel, bitmap.GetPixel(x + 1, y + 2)) > diffTreshold)
                    {
                        if (goodX < 0)
                        {
                            goodX = x;
                            goodY = y;
                        }
                        if (false
                            || (true
                                && ColorDifference(bitmap.GetPixel(x + 2, y + 1), bitmap.GetPixel(x + 2, y + 2)) > diffTreshold
                                && ColorDifference(bitmap.GetPixel(x + 2, y + 1), bitmap.GetPixel(x + 2, y)) > diffTreshold
                                && ColorDifference(bitmap.GetPixel(x + 2, y + 2), bitmap.GetPixel(x, y + 2)) > diffTreshold)
                            || (true
                                && ColorDifference(bitmap.GetPixel(x + 1, y + 2), bitmap.GetPixel(x, y + 2)) > diffTreshold
                                && ColorDifference(bitmap.GetPixel(x + 1, y + 2), bitmap.GetPixel(x + 2, y + 2)) > diffTreshold
                                && ColorDifference(bitmap.GetPixel(x + 2, y + 2), bitmap.GetPixel(x + 2, y)) > diffTreshold)
                            || (true
                                && ColorDifference(bitmap.GetPixel(x + 2, y + 1), bitmap.GetPixel(x + 2, y)) > diffTreshold
                                && ColorDifference(bitmap.GetPixel(x + 1, y + 2), bitmap.GetPixel(x, y + 2)) > diffTreshold))
                        {
                            bestX = x;
                            bestY = y;
                            break;
                        }
                    }
                }
                if (bestX >= 0)
                    break;
            }
            if (bestX < 0)
            {
                bestX = goodX;
                bestY = goodY;
            }

            bestX = Math.Max(0, Math.Min(bestX, bitmap.Width - fragmentWidth - 1));
            bestY = Math.Max(0, Math.Min(bestY, bitmap.Height - fragmentHeight - 1));
            Point bestPoint = new(bestX, bestY);
            return bestPoint;
        }

        private void FocusClipText()
        {
            if (htmlMode)
                htmlTextBox.Document.Focus();
            else //if (richTextBox.Enabled)
                richTextBox.Focus();
        }

        //private string FormattedClipNumericPropery(int number, string unit)
        //{
        //    NumberFormatInfo numberFormat = new CultureInfo(Locale).NumberFormat;
        //    numberFormat.NumberDecimalDigits = 0;
        //    numberFormat.NumberGroupSeparator = " ";
        //    return number.ToString("N", numberFormat) + " " + unit;
        //}

        private Bitmap getBitmapFromUrl(string imageUrl)
        {
            //Bitmap bitmap = new Bitmap();
            Bitmap bitmap = null;
            if (settings.AllowDownloadThumbnail || imageUrl.StartsWith("data:image") || File.Exists(imageUrl))
                using (System.Net.Http.HttpClient webClient = new())
                {
                    try
                    {
                        var tempBuffer = webClient.GetByteArrayAsync(imageUrl).Result;
                        using var ms = new MemoryStream(tempBuffer);
                        bitmap = new Bitmap(ms);
                    }
                    catch
                    {
                    }
                }
            return bitmap;
        }

        private string GetClipTempFile(out string fileEditor, Clip rowReader = null)
        {
            fileEditor = "";
            if (rowReader == null)
                rowReader = clip;
            if (rowReader == null)
                return "";
            string type = rowReader.Type;
            //string TempFile = Path.GetTempFileName();
            string tempFile = clipTempFile(rowReader);
            if (tempFile == "")
            {
                MessageBox.Show(this, Properties.Resources.ClipFileAlreadyOpened);
                return "";
            }
            if (type == "text" /*|| type == "file"*/)
            {
                File.WriteAllText(tempFile, rowReader.Text, Encoding.UTF8);
                fileEditor = settings.TextEditor;
            }
            else if (type == "rtf")
            {
                RichTextBox rtb = new()
                {
                    Rtf = rowReader.RichText
                };
                rtb.SaveFile(tempFile);
                fileEditor = settings.RtfEditor;
            }
            else if (type == "html")
            {
                File.WriteAllText(tempFile, GetHtmlFromHtmlClipText(true), Encoding.UTF8);
                fileEditor = settings.HtmlEditor;
            }
            else if (type == "img")
            {
                Image image = GetImageFromBinary((byte[])rowReader.Binary);
                image.Save(tempFile);
                fileEditor = settings.ImageEditor;
            }
            else if (type == "file")
            {
                var tokens = TextToLines(rowReader.Text);
                tempFile = tokens[0];
                if (!File.Exists(tempFile))
                    tempFile = "";
            }
            return tempFile;
        }

        private mshtml.IHTMLTxtRange GetHtmlCurrentTextRangeOrAllDocument(bool onlySelection = false)
        {
            mshtml.IHTMLDocument2 htmlDoc = (mshtml.IHTMLDocument2)htmlTextBox.Document.DomDocument;
            mshtml.IHTMLBodyElement body = htmlDoc.body as mshtml.IHTMLBodyElement;
            mshtml.IHTMLSelectionObject sel = htmlDoc.selection;
            //sel.empty(); // get an empty selection, so we start from the beginning
            mshtml.IHTMLTxtRange range = null;
            try
            {
                range = (mshtml.IHTMLTxtRange)sel.createRange();
            }
            catch { }
            if (range == null && !onlySelection)
                range = body.createTextRange();
            return range;
        }

        private string GetHtmlFromHtmlClipText(bool AddSourceUrlComment = false)
        {
            string htmlClipText = clip.HtmlText;
            if (String.IsNullOrEmpty(htmlClipText))
                return "";
            int indexOfHtlTag = htmlClipText.IndexOf("<html", StringComparison.OrdinalIgnoreCase);
            if (indexOfHtlTag < 0)
                return "";
            string result = htmlClipText.Substring(indexOfHtlTag);
            if (AddSourceUrlComment)
                result = result + "\n<!-- Original URL - " + clip.Url + " -->";
            return result;
        }

        private int GetHtmlPosition(out int length)
        {
            length = 0;
            int start = 0;
            mshtml.IHTMLDocument2 htmlDoc = htmlTextBox.Document.DomDocument as mshtml.IHTMLDocument2;
            mshtml.IHTMLTxtRange range = null;
            try
            {
                range = (IHTMLTxtRange)htmlDoc.selection.createRange();
            }
            catch
            {
            }
            if (range == null)
                return start;
            if (!String.IsNullOrEmpty(range.text))
                length = range.text.Length
                    //- GetNormalizedTextDeltaSize(range.text)
                    ;
            string innerText = htmlDoc.body.innerText;
            if (String.IsNullOrEmpty(innerText) || innerText.Length > 3000)
                // Long html will make moveStart slow
                return start;
            range.collapse();
            range.moveStart("character", -100000);
            if (!String.IsNullOrEmpty(range.text))
                start = range.text.Length
                    //- GetNormalizedTextDeltaSize(range.text)
                    ;
            if (!String.IsNullOrEmpty(innerText))
            {
                int maxStart = innerText.Length
                    //- GetNormalizedTextDeltaSize(innerText)
                    ;
                if (start > maxStart)
                    start = maxStart;
            }
            return start;
        }

        private string getSelectedOrAllText()
        {
            string selectedText = richTextBox.SelectedText;
            if (selectedText == "")
                selectedText = clip.Text;
            return selectedText;
        }

        private string GetSelectedTextOfClip(bool onlySelectedPlainText = true)
        {
            string selectedText = "";
            mshtml.IHTMLTxtRange htmlSelection = null;
            if (clip.Type == "html")
                htmlSelection = GetHtmlCurrentTextRangeOrAllDocument(true);
            bool selectedPlainTextMode = true
                                         && onlySelectedPlainText
                                         && (false
                                             || !String.IsNullOrEmpty(richTextBox.SelectedText)
                                             || htmlSelection != null && !String.IsNullOrEmpty(htmlSelection.text));
            if (selectedPlainTextMode && !String.IsNullOrEmpty(richTextBox.SelectedText))
            {
                selectedText = richTextBox.SelectedText;
            }
            else if (selectedPlainTextMode && !String.IsNullOrEmpty(htmlSelection.text))
            {
                selectedText = htmlSelection.text;
            }
            else if (EditMode)
                selectedText = richTextBox.Text;
            return selectedText;
        }

        private void GotoLastRow(bool keepTextSelectionIfIDChanged = false)
        {
            if (dataGridView.Rows.Count > 0)
            {
                allowRowLoad = false;
                clipBindingSource.MoveFirst(); // It changes selected row
                allowRowLoad = true;
                if (dataGridView.CurrentRow != null)
                    SelectCurrentRow(false, true, true, keepTextSelectionIfIDChanged);
            }
            LoadClipIfChangedID(false, true, keepTextSelectionIfIDChanged);
        }

        private void GotoSearchMatchInList(bool Forward = true, bool resetPosition = false)
        {
            if (searchMatchedIDs.Count == 0)
                return;
            int currentListMatchIndex;
            if (Forward)
            {
                if (resetPosition)
                    currentListMatchIndex = 0;
                else
                {
                    currentListMatchIndex = searchMatchedIDs.Count - 1;
                    if (dataGridView.CurrentRow != null)
                    {
                        int ListMatchIndex;
                        for (int index = dataGridView.CurrentRow.Index + 1; index < dataGridView.RowCount; index++)
                        {
                            DataRowView dataRow = (DataRowView)dataGridView.Rows[index].DataBoundItem;
                            ListMatchIndex = searchMatchedIDs.IndexOf((int)dataRow["Id"]);
                            if (ListMatchIndex > -1)
                            {
                                currentListMatchIndex = ListMatchIndex;
                                break;
                            }
                        }
                    }
                }
            }
            else
            {
                currentListMatchIndex = 0;
                if (dataGridView.CurrentRow != null)
                {
                    int ListMatchIndex;
                    for (int index = dataGridView.CurrentRow.Index - 1; index >= 0; index--)
                    {
                        DataRowView dataRow = (DataRowView)dataGridView.Rows[index].DataBoundItem;
                        ListMatchIndex = searchMatchedIDs.IndexOf((int)dataRow["Id"]);
                        if (ListMatchIndex > -1)
                        {
                            currentListMatchIndex = ListMatchIndex;
                            break;
                        }
                    }
                }
            }
            RestoreSelectedCurrentClip(true, searchMatchedIDs[currentListMatchIndex]);
        }

        private void hook_KeyPressed(object sender, KeyPressedEventArgs e)
        {
            if (!AllowHotkeyProcess)
                return;
            string hotkeyTitle = KeyboardHook.HotkeyTitle(e.Key, e.Modifier);
            if (hotkeyTitle == settings.GlobalHotkeyOpenLast)
            {
                if (IsVisible() && this.ContainsFocus && MarkFilter.SelectedValue.ToString() != "favorite") // Sometimes it can cotain focus but be not visible!
                    this.Close();
                else
                {
                    ShowForPaste(false, true);
                    dataGridView.Focus();
                }
            }
            else if (hotkeyTitle == settings.GlobalHotkeyOpenCurrent)
            {
                if (IsVisible() && this.ContainsFocus)
                    this.Close();
                else
                {
                    ShowForPaste();
                    //dataGridView.Focus();
                }
            }
            else if (hotkeyTitle == settings.GlobalHotkeyOpenFavorites)
            {
                if (IsVisible() && this.ContainsFocus && MarkFilter.SelectedValue.ToString() == "favorite")
                    this.Close();
                else
                {
                    ShowForPaste(true, true);
                    dataGridView.Focus();
                }
            }
            else if (hotkeyTitle == settings.GlobalHotkeyIncrementalPaste)
            {
                PasteAndSelectNext(1);
            }
            else if (hotkeyTitle == settings.GlobalHotkeyDecrementalPaste)
            {
                PasteAndSelectNext(-1);
            }
            else if (hotkeyTitle == settings.GlobalHotkeyCompareLastClips)
            {
                if (filterOn)
                    ClearFilter(-1, false, true);
                toolStripMenuItemCompareLastClips_Click();
            }
            else if (hotkeyTitle == settings.GlobalHotkeyPasteText)
            {
                if (filterOn)
                    ClearFilter(-1, false, true);
                SendPasteClipExpress(dataGridView.Rows[0], PasteMethod.Text);
            }
            else if (hotkeyTitle == settings.GlobalHotkeySimulateInput)
            {
                if (filterOn)
                    ClearFilter(-1, false, true);
                SendPasteClipExpress(dataGridView.Rows[0], PasteMethod.SendCharsFast);
            }
            else if (hotkeyTitle == settings.GlobalHotkeySwitchMonitoring)
            {
                SwitchMonitoringClipboard(true);
            }
            else if (hotkeyTitle == settings.GlobalHotkeyForcedCapture)
            {
                if (!settings.MonitoringClipboard)
                {
                    settings.MonitoringClipboard = true;
                    ConnectClipboard();
                    tempCaptureTimer.Interval = 2000;
                    tempCaptureTimer.Start();
                }
                Paster.SendCopy(false);
            }
        }

        private void PasteAndSelectNext(int direction)
        {
            AllowHotkeyProcess = false;
            try
            {
                SendPasteClipExpress(null, PasteMethod.Standard, false, true);
                // https://www.hostedredmine.com/issues/925182
                //if ((e.Modifier & EnumModifierKeys.Alt) != 0)
                //    keybd_event((byte) VirtualKeyCode.MENU, 0x38, 0, 0); // LEFT
                //if ((e.Modifier & EnumModifierKeys.Control) != 0)
                //    keybd_event((byte) VirtualKeyCode.CONTROL, 0x1D, 0, 0);
                //if ((e.Modifier & EnumModifierKeys.Shift) != 0)
                //    keybd_event((byte) VirtualKeyCode.SHIFT, 0x2A, 0, 0);
                DataRow oldCurrentDataRow = ((DataRowView)clipBindingSource.Current).Row;
                if (direction > 0)
                    clipBindingSource.MovePrevious();
                else
                    clipBindingSource.MoveNext();
                DataRow CurrentDataRow = ((DataRowView)clipBindingSource.Current).Row;
                notifyIcon.Visible = true;
                string messageText;
                if (oldCurrentDataRow == CurrentDataRow)
                    messageText = Properties.Resources.PastedLastClip;
                else
                    messageText = CurrentDataRow["Title"].ToString();
                notifyIcon.ShowBalloonTip(3000, Properties.Resources.NextClip, messageText, ToolTipIcon.Info);
            }
            catch (Exception)
            {
                int dummy = 1;
            }
            finally
            {
                AllowHotkeyProcess = true;
            }
        }


        private void htmlMenuItemCopy_Click(object sender, EventArgs e)
        {
            htmlTextBox.Document.ExecCommand("COPY", false, 0);
        }

        private bool htmlTextBoxDocumentClick(mshtml.IHTMLEventObj e)
        {
            if (e.altKey)
            {
                IHTMLElement hlink = e.srcElement;
                openLinkInBrowserToolStripMenuItem_Click();
            }
            return false;
        }

        private void htmlTextBoxDocumentKeyDown(Object sender, HtmlElementEventArgs e)
        {
            e.ReturnValue = false
                            || e.CtrlKeyPressed
                            || e.AltKeyPressed
                            || e.KeyPressedCode == (int)Keys.Down
                            || e.KeyPressedCode == (int)Keys.Up
                            || e.KeyPressedCode == (int)Keys.Left
                            || e.KeyPressedCode == (int)Keys.Right
                            || e.KeyPressedCode == (int)Keys.PageDown
                            || e.KeyPressedCode == (int)Keys.PageUp
                            || e.KeyPressedCode == (int)Keys.Home
                            || e.KeyPressedCode == (int)Keys.End
                ;
            if (e.KeyPressedCode == (int)Keys.Escape)
                Close();
            else if (e.KeyPressedCode == (int)Keys.Enter)
            {
                ProcessEnterKeyDown(e.CtrlKeyPressed, e.ShiftKeyPressed);
            }
        }

        private void htmlTextBoxDocumentSelectionChange(Object sender = null, EventArgs e = null)
        {
            if (!allowTextPositionChangeUpdate)
                return;
            selectionStart = GetHtmlPosition(out selectionLength);
            UpdateClipContentPositionIndicator(0, 0);
        }

        private void htmlTextBoxDrag(Object sender, HtmlElementEventArgs e)
        {
            e.ReturnValue = false;
        }

        private void htmlTextBoxMouseDown(IHTMLEventObj pEvtObj)
        {
            lastClickedHtmlElement = htmlTextBox.Document.GetElementFromPoint(htmlTextBox.PointToClient(MousePosition));
            bool isLink = (String.Compare(lastClickedHtmlElement.TagName, "A", true) == 0);
            htmlMenuItemCopyLinkAdress.Enabled = isLink;
            htmlMenuItemOpenLink.Enabled = isLink;
        }

        private void ImageControl_DoubleClick(object sender, EventArgs e)
        {
            OpenClipFile();
        }

        private void ImageControl_MouseWheel(object sender, MouseEventArgs e)
        {
            UpdateClipContentPositionIndicator();
        }

        private void ImageControl_Resize(object sender, EventArgs e)
        {
            UpdateClipContentPositionIndicator();
        }

        private void ImageControl_ZoomChanged(object sender, EventArgs e)
        {
            UpdateClipContentPositionIndicator();
        }

        private void importToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = false;
            openFileDialog.Filter = "Clippy clips|*.cac|All|*.*";
            if (openFileDialog.ShowDialog(this) != DialogResult.OK)
                return;
            DataTable dataTable = new DataTable();
            dataTable.Locale = CultureInfo.InvariantCulture;
            try
            {
                dataTable.ReadXml(openFileDialog.FileName);
            }
            catch
            {
                MessageBox.Show(this, Properties.Resources.WrongFileFormat, Application.ProductName);
                return;
            }
            foreach (DataRow importedRow in dataTable.Rows)
            {
                AddClip((byte[])importedRow["Binary"], (byte[])importedRow["ImageSample"], importedRow["HtmlText"].ToString(), importedRow["RichText"].ToString(), importedRow["Type"].ToString(), importedRow["Text"].ToString(),
                    importedRow["application"].ToString(), importedRow["window"].ToString(), importedRow["url"].ToString(), Convert.ToInt32(importedRow["chars"].ToString()),
                    importedRow["AppPath"].ToString(), false, Convert.ToBoolean(importedRow["Favorite"].ToString()), false, importedRow["title"].ToString(), DateTime.Parse(importedRow["Created"].ToString()));
            }
            ReloadList(false, 0, false, null, true);
        }

        private bool IsTextType(string type = "")
        {
            if (type == "")
                type = (string)clip.Type;
            return type == "rtf" || type == "text" || type == "html";
        }

        private bool IsVisible()
        {
            return this.Visible && this.Top > maxWindowCoordForHiddenState;
        }

        private string JoinOrPasteTextOfClips(PasteMethod itemPasteMethod, out int count, string DelimiterForTextJoin = null)
        {
            string agregateTextToPaste = "";
            bool pasteDelimiter = false;
            count = dataGridView.SelectedRows.Count;
            for (int i = count - 1; i >= 0; i--)
            {
                DataGridViewRow selectedRow = dataGridView.SelectedRows[i];
                agregateTextToPaste += SendPasteClipExpress(selectedRow, itemPasteMethod, pasteDelimiter, false, DelimiterForTextJoin);
                pasteDelimiter = true;
            }
            return agregateTextToPaste;
        }

        private void LoadClipIfChangedID(bool forceRowLoad = false, bool keepTextSelection = true,  bool keepTextSelectionIfIDChanged = false)
        {
            bool currentIDChanged = CurrentIDChanged();
            if (forceRowLoad || currentIDChanged)
            {
                if (currentIDChanged)
                {
                    if (!keepTextSelectionIfIDChanged)
                        keepTextSelection = false;
                    EditMode = false;
                    UpdateControlsStates();
                }
                int NewSelectionStart, NewSelectionLength;
                if (keepTextSelection)
                {
                    NewSelectionStart = -1;
                    NewSelectionLength = -1;
                }
                else
                {
                    NewSelectionStart = 0;
                    NewSelectionLength = 0;
                }
                AfterRowLoad(false, -1, NewSelectionStart, NewSelectionLength);
            }
        }

        private void LoadDataboundItems(int firstDisplayedRowIndex, int lastDisplayedRowIndex)
        {
            int bufferSize = 50;
            int firstLoadedRowIndex = Math.Max(firstDisplayedRowIndex - bufferSize, 0);
            int lastLoadedRowIndex = Math.Min(lastDisplayedRowIndex + bufferSize, dataGridView.RowCount - 1);

            //DataTable table = (DataTable)clipBindingSource.DataSource;
            for (int rowIndex = firstLoadedRowIndex; rowIndex <= lastLoadedRowIndex; rowIndex++)
            {
                DataRowView drv = (dataGridView.Rows[rowIndex].DataBoundItem as DataRowView);
                if (!String.IsNullOrEmpty(drv["type"].ToString()))
                    continue;

                int rowId = (int)drv["id"];
                var clip_get = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Id == rowId);
                if (clip_get != null)
                {
                    DataRowView row = (dataGridView.Rows[rowIndex].DataBoundItem as DataRowView);
                    row["Used"] = clip_get.Used;
                    row["Title"] = clip_get.Title;
                    row["Chars"] = clip_get.Chars;
                    row["Type"] = clip_get.Type; 
                    row["Favorite"] = !clip_get.Favorite && clip_get.Favorite;
                    row["ImageSample"] = clip_get.ImageSample;
                    row["AppPath"] = clip_get.AppPath;
                    row["Size"] = clip_get.Size;
                    row["Created"] = clip_get.Created;

                }
            }

        }

        private void LoadSettings()
        {
            UpdateControlsStates();
            this.SuspendLayout();
            //UpdateCurrentCulture();
            //cultureManager1.UICulture = Thread.CurrentThread.CurrentUICulture;

            UpdateWindowTitle(true);
            notifyIcon.Icon = Properties.Resources.clipboard;

            UpdateIgnoreModulesInClipCapture();
            BindingList<ListItemNameText> comboItemsTypes = (BindingList<ListItemNameText>)TypeFilter.DataSource;
            foreach (ListItemNameText item in comboItemsTypes)
            {
                //item.Text = CurrentLangResourceManager.GetString(item.Name);
                if (item.Name.StartsWith("text_"))
                {
                    item.Text = item.Name.Replace("text_", "");
                }
                else
                {
                    item.Text = item.Name;
                }
            }
            // To refresh text in list
            TypeFilter.DisplayMember = "";
            TypeFilter.DisplayMember = "Text";

            
            //VisibleUserSettings allSettings = new VisibleUserSettings(this);
            searchAllFieldsMenuItem.ToolTipText = Properties.Resources.SearchAllFields.ToString(); 
            toolStripButtonAutoSelectMatch.ToolTipText = Properties.Resources.AutoSelectMatch.ToString();
            autoselectMatchedClipMenuItem.ToolTipText = Properties.Resources.AutoSelectMatchedClip.ToString();
            filterListBySearchStringMenuItem.ToolTipText = Properties.Resources.FilterListBySearchString.ToString();
            toolStripMenuItemSearchCaseSensitive.ToolTipText = Properties.Resources.SearchCaseSensitive.ToString();
            ignoreBigTextsToolStripMenuItem.ToolTipText = Properties.Resources.SearchIgnoreBigTexts.ToString();
            toolStripMenuItemSearchWordsIndependently.ToolTipText = Properties.Resources.SearchWordsIndependently.ToString();
            toolStripMenuItemSearchWildcards.ToolTipText = Properties.Resources.SearchWildcards.ToString();
            //moveCopiedClipToTopToolStripButton.ToolTipText = Properties.Resources.MoveCopiedClipToTop.ToString();
            moveCopiedClipToTopToolStripMenuItem.ToolTipText = Properties.Resources.MoveCopiedClipToTop.ToString();
            textFormattingToolStripMenuItem.ToolTipText = Properties.Resources.ShowNativeTextFormatting.ToString();
           // toolStripButtonTextFormatting.ToolTipText = Properties.Resources.ShowNativeTextFormatting.ToString();
            //toolStripButtonMonospacedFont.ToolTipText = Properties.Resources.MonospacedFont.ToString();
            monospacedFontToolStripMenuItem.ToolTipText = Properties.Resources.MonospacedFont.ToString();
            wordWrapToolStripMenuItem.ToolTipText = Properties.Resources.WordWrap.ToString();
            //toolStripButtonWordWrap.ToolTipText = Properties.Resources.WordWrap.ToString();
            //toolStripButtonSecondaryColumns.ToolTipText = Properties.Resources.ShowSecondaryColumns.ToString();
            toolStripMenuItemSecondaryColumns.ToolTipText = Properties.Resources.ShowSecondaryColumns.ToString();

            BindingList<ListItemNameText> comboItemsMarks = (BindingList<ListItemNameText>)MarkFilter.DataSource;
            foreach (ListItemNameText item in comboItemsMarks)
            {
                item.Text = item.Name;
            }
            // To refresh text in list
            MarkFilter.DisplayMember = "";
            MarkFilter.DisplayMember = "Text";
            settings.RestoreCaretPositionOnFocusReturn = false; // disabled
            dataGridView.RowsDefaultCellStyle.Font = new Font(settings.dataGridViewFontFamily, settings.dataGridViewFontSize);
            dataGridView.Columns["ColumnCreated"].DefaultCellStyle.Format = "HH:mmm:ss dd.MM";
            //dataGridView.Columns["VisualWeight"].Width = (int)dataGridView.RowsDefaultCellStyle.Font.Size;
            dataGridView.RowTemplate.Height = (int)(dataGridView.RowsDefaultCellStyle.Font.Size + 14);

            UpdateColumnsSet();
            AfterRowLoad();
            this.ResumeLayout();
        }

        private void LoadVisibleRows()
        {
            int visibleRowsCount = dataGridView.DisplayedRowCount(true);
            int firstDisplayedRowIndex = dataGridView.FirstDisplayedCell.RowIndex;
            int lastDisplayedRowIndex = Math.Min(firstDisplayedRowIndex + visibleRowsCount, dataGridView.RowCount - 1);
            bool needLoad = false;
            for (int rowIndex = firstDisplayedRowIndex; rowIndex <= lastDisplayedRowIndex; rowIndex++)
            {
                DataRowView drv = (dataGridView.Rows[rowIndex].DataBoundItem as DataRowView);
                if (String.IsNullOrEmpty(drv["type"].ToString()))
                {
                    needLoad = true;
                    break;
                }
            }
            if (!needLoad)
                return;
            LoadDataboundItems(firstDisplayedRowIndex, lastDisplayedRowIndex);
        }

        private void Main_Activated(object sender, EventArgs e)
        {
            if (settings.FastWindowOpen)
            {
                RestoreWindowIfMinimized();
            }
            SetForegroundWindow(this.Handle);
            TimeFromWindowOpen = DateTime.Now;
        }

        private void Main_Deactivate(object sender, EventArgs e)
        {
            if (settings.FastWindowOpen)
            {
                if (this.Top > maxWindowCoordForHiddenState)
                    factualTop = this.Top;
            }
        }

        private void Main_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Visible = false;
            if (this.Left == -32000)
                settings.WindowPositionX = this.RestoreBounds.Left;
            else
                settings.WindowPositionX = this.Left;
            if (this.Top == -32000)
                settings.WindowPositionY = this.RestoreBounds.Top;
            else if (this.Top == maxWindowCoordForHiddenState)
                settings.WindowPositionY = factualTop;
            else
                settings.WindowPositionY = this.Top;
            settings.dataGridViewWidth = splitContainer1.SplitterDistance;

            //settings.Save(); // Not all properties were saved here. For example ShowInTaskbar was not saved
            RemoveClipboardFormatListener(this.Handle);
            UnhookWinEvent(HookChangeActiveWindow);

            if (settings.DeleteNonFavoriteClipsOnExit)
                deleteAllNonFavoriteClips();


            //if (oneDB == false && defaultDb != null)
            //{
            //    var col = defaultDb.GetCollection<Settings>("settings").Update(settings);
            //    defaultDb.Dispose();
            //}
            //else
            //{
                liteDb.GetCollection<Settings>("settings").Update(settings);
                liteDb.Dispose();
            //}

            GC.Collect();
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                bool lastActSet = false;

                lastActSet = SetForegroundWindow(lastActiveParentWindow);

                if (!lastActSet)
                    SetActiveWindow(IntPtr.Zero);
                if (settings.FastWindowOpen)
                {
                    this.Top = maxWindowCoordForHiddenState;
                }
                else
                {
                    this.SuspendLayout();
                    this.FormBorderStyle = FormBorderStyle.FixedToolWindow; // To disable animation
                    this.Hide();
                    this.ResumeLayout();
                }
                e.Cancel = true;
            }
            else
            {
                if (WindowState == FormWindowState.Normal)
                {
                    settings.MainWindowSizeWidth = Size.Width;
                    settings.MainWindowSizeHeight = Size.Height;
                }
                else
                {
                    settings.MainWindowSizeWidth = RestoreBounds.Size.Width;
                    settings.MainWindowSizeHeight = RestoreBounds.Size.Height;
                }

            }
        }

        private void Main_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                e.Handled = true;
                this.Close();
            }
            if (e.KeyCode == Keys.Enter)
            {
                if (ProcessEnterKeyDown(e.Control, e.Shift))
                    return;
                e.Handled = true;
            }
            if (e.KeyCode == Keys.Tab)
            {
                e.Handled = true;
                FocusClipText();
            }
            if (true
                && (DateTime.Now - TimeFromWindowOpen).TotalMilliseconds < 1000 // Temporary block main menu activation to avoid unwanted action while opening main window with ALT+* hotkey
                && (e.Modifiers == Keys.Alt)
                && (e.KeyCode == Keys.Menu))
            {
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
        }

        private void MarkFilter_SelectedValueChanged(object sender, EventArgs e)
        {
            if (AllowFilterProcessing)
            {
                ReloadList();
            }
        }

        private void MarkLinksInRichTextBox(RichTextBox control, out MatchCollection matches)
        {
            MarkRegExpMatchesInRichTextBox(control, "(" + fileOrFolderPattern + "|" + LinkPattern + ")", Color.Blue, false, true, false, out matches);
        }

        private void MarkRegExpMatchesInRichTextBox(RichTextBox control, string pattern, Color color, bool allowDymanicColor, bool underline, bool bold, out MatchCollection matches)
        {
            RegexOptions options = RegexOptions.Singleline;
            if (!settings.SearchCaseSensitive)
                options = options | RegexOptions.IgnoreCase;
            matches = Regex.Matches(control.Text, pattern, options);
            control.DeselectAll();
            int maxMarked = 50; // prevent slow down
            foreach (Match match in matches)
            {
                int startGroup = 2;
                if (match.Groups.Count < 2)
                    throw new ArgumentNullException("Wrong regexp pattern");
                control.SelectionStart = match.Groups[1].Index;
                control.SelectionLength = match.Groups[1].Length;
                if (allowDymanicColor && match.Groups.Count > 3)
                {
                    for (int counter = startGroup; counter < match.Groups.Count; counter++)
                    {
                        if (match.Groups[counter].Success)
                        {
                            color = _wordColors[(counter - startGroup) % _wordColors.Length];
                            break;
                        }
                    }
                }
                control.SelectionColor = color;
                if (control.SelectionFont != null)
                {
                    FontStyle newstyle = control.SelectionFont.Style;
                    if (bold)
                        newstyle = newstyle | FontStyle.Bold;
                    if (underline)
                        newstyle = newstyle | FontStyle.Underline;
                    if (newstyle != control.SelectionFont.Style)
                        control.SelectionFont = new Font(control.SelectionFont, newstyle);
                }
                maxMarked--;
                if (maxMarked < 0)
                    break;
            }
            control.DeselectAll();
            control.SelectionColor = new Color();
            control.SelectionFont = new Font(control.SelectionFont, FontStyle.Regular);
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            monthCalendar1.Hide();
            periodFilterOn = true;
            ReloadList();

            // Turn on secodnary columns
            if (!settings.ShowSecondaryColumns)
                toolStripMenuItemSecondaryColumns_Click();
        }

        private void monthCalendar1_Leave(object sender, EventArgs e)
        {
            monthCalendar1.Hide();
        }

        private void MoveSelectedRows(int shiftType)
        {
            // shiftType - 0 - TOP
            //            -1 - UP
            //             1 - DOWN
            if (dataGridView.CurrentRow == null)
                return;
            int newID;
            int newCurrentID = 0;
            int currentRowIndex = dataGridView.CurrentRow.Index;
            List<int> seletedRowIndexes = new List<int>();
            int counter = 0;
            Dictionary<int, int> order = new Dictionary<int, int>();
            foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
            {
                seletedRowIndexes.Add(selectedRow.Index);
                order.Add(counter, selectedRow.Index);
                counter++;
            }
            IOrderedEnumerable<int> SortedRowIndexes;
            if (shiftType < 0)
                SortedRowIndexes = seletedRowIndexes.OrderBy(i => i);
            else
                SortedRowIndexes = seletedRowIndexes.OrderByDescending(i => i);
            foreach (int selectedRowIndex in SortedRowIndexes.ToList())
            {
                DataRow selectedDataRow = ((DataRowView)clipBindingSource[selectedRowIndex]).Row;
                int oldID = (int)selectedDataRow["ID"];
                if (shiftType != 0)
                {
                    if (selectedRowIndex + shiftType < 0 || selectedRowIndex + shiftType > clipBindingSource.Count - 1)
                        continue;
                    DataRow exchangeDataRow = ((DataRowView)clipBindingSource[selectedRowIndex + shiftType]).Row;
                    newID = (int)exchangeDataRow["ID"];
                }
                else
                {
                    LastId++;
                    newID = LastId;
                }
                if (currentRowIndex == selectedRowIndex)
                    newCurrentID = newID;
                int tempID = LastId + 1;
                //string sql = "Update Clips set Id=@NewId where Id=@Id";
                //SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                if (newID != tempID)
                {
                    //command.Parameters.AddWithValue("@Id", newID);
                    //command.Parameters.AddWithValue("@NewID", tempID);
                    //command.ExecuteNonQuery();

                    var clip_get1 = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Id==newID);
                    if (clip_get1 != null)
                    {
                        clip_get1.Id = tempID;
                        var clip_upd = liteDb.GetCollection<Clip>("clips").Update(clip_get1);
                    }
                    
                }
                //command.Parameters.AddWithValue("@Id", oldID);
                //command.Parameters.AddWithValue("@NewID", newID);
                //command.ExecuteNonQuery();

                var clip_get2 = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Id==oldID);
                if (clip_get2 != null)
                {
                    clip_get2.Id = newID;
                    var clip_upd = liteDb.GetCollection<Clip>("clips").Update(clip_get2);
                }

                if (newID != tempID)
                {
                    //command.Parameters.AddWithValue("@Id", tempID);
                    //command.Parameters.AddWithValue("@NewID", oldID);
                    //command.ExecuteNonQuery();

                    var clip_get3 = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Id==tempID);
                    if (clip_get3 != null)
                    {
                        clip_get3.Id = oldID;
                        var clip_upd = liteDb.GetCollection<Clip>("clips").Update(clip_get3);
                    }
                }
                order[order.FirstOrDefault(x => x.Value == selectedRowIndex).Key] = newID;
                RegisterClipIdChange(oldID, newID);
            }
            List<int> ids = order.Select(d => d.Value).ToList();
            ids.Reverse();
            ReloadList(false, newCurrentID, true, ids, true);
        }

        private void NextMatchListMenuItem_Click(object sender, EventArgs e)
        {
            GotoSearchMatchInList();
        }

        private void notifyIcon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (IsVisible())
                    this.Close();
                else
                    ShowForPaste(false, false, true);
            }
        }

        private void OnClipContentSelectionChange()
        {
            if (htmlMode)
                htmlTextBoxDocumentSelectionChange();
            else
                richTextBox_SelectionChanged();
        }

        private void OpenClipFile(bool defaultAppMode = true)
        {
            string tempFile = GetClipTempFile(out string fileEditor);
            if (String.IsNullOrEmpty(tempFile))
                return;
            try
            {
                if (defaultAppMode)
                {
                    string command;
                    string argument = "";
                    if (!String.IsNullOrEmpty(fileEditor))
                    {
                        command = fileEditor;
                        argument = "\"" + tempFile + "\"";
                    }
                    else
                    {
                        command = tempFile;
                    }
                    Process.Start(command, argument);
                }
                else
                {
                    // http://stackoverflow.com/questions/4726441/how-can-i-show-the-open-with-file-dialog

                    // Not reliable
                    var args = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "shell32.dll");
                    args += ",OpenAs_RunDLL \"" + tempFile + "\"";
                    Process.Start("rundll32.exe", args);
                }
            }
            catch
            {
                // ignored
            }
        }

        private bool OpenLinkFromTextBox(MatchCollection matches, int SelectionStart, bool allowOpenUnknown = true)
        {
            // Result - true if known link found
            foreach (Match match in matches)
            {
                if (match.Index <= SelectionStart && (match.Index + match.Length) >= SelectionStart)
                {
                    int startIndex1C = 2;
                    if (match.Groups[startIndex1C + 5].Success) // File link
                    {
                        string filePath = match.Value;
                        if (!File.Exists(filePath) && !Directory.Exists(filePath))
                        {
                            return true;
                        }
                        string argument = "/select, \"" + filePath + "\"";
                        System.Diagnostics.Process.Start("explorer.exe", argument);
                    }

                    if (allowOpenUnknown)
                    {
                        try
                        {
                            Process.Start(match.Value);
                        }
                        catch
                        {
                            // for example file://C:/Users/Donny/AppData/Local/Temp/Clip.html
                        }
                    }
                    return true;
                }
            }
            return false;
        }

        private void OpenLinkIfAltPressed(RichTextBox sender, EventArgs e, MatchCollection matches, bool checkAlt = true)
        {
            Keys mod = Control.ModifierKeys & Keys.Modifiers;
            bool altOnly = mod == Keys.Alt;
            if (!checkAlt || altOnly)
                OpenLinkFromTextBox(matches, sender.SelectionStart);
        }

        private void PassKeyToGrid(bool downOrUp, KeyEventArgs e)
        {
            if (IsKeyPassedFromFilterToGrid(e.KeyCode, e.Control))
            {
                sendKey(dataGridView.Handle, e.KeyCode, false, downOrUp, !downOrUp);
                e.Handled = true;
            }
        }

        private void PrepareRow(DataGridViewRow row = null)
        {
            LoadVisibleRows();
            if (row == null)
                row = dataGridView.CurrentRow;
            DataRowView dataRowView = (DataRowView)row.DataBoundItem;
            int shortSize = dataRowView.Row["Chars"].ToString().Length;
            if (shortSize > 2)
                row.Cells["VisualWeight"].Value = shortSize;
            string clipType = dataRowView.Row["Type"].ToString();
            Bitmap image = null;
            switch (clipType)
            {
                case "text":
                    image = imageText;
                    break;

                case "html":
                    image = imageHtml;
                    break;

                case "rtf":
                    image = imageRtf;
                    break;

                case "file":
                    image = imageFile;
                    break;

                case "img":
                    image = imageImg;
                    break;

                default:
                    break;
            }
            if (image != null)
            {
                row.Cells["TypeImage"].Value = image;
            }
            row.Cells["ColumnTitle"].Value = dataRowView.Row["Title"].ToString();

            string textPattern = RegexpPatternFromTextFilter();
            _richTextBox.Clear();
            _richTextBox.Font = dataGridView.RowsDefaultCellStyle.Font;
            _richTextBox.Text = row.Cells["ColumnTitle"].Value.ToString();
            if (!String.IsNullOrEmpty(textPattern))
            {
                MatchCollection tempMatches;
                MarkRegExpMatchesInRichTextBox(_richTextBox, textPattern, Color.Red, true, false, true, out tempMatches);
            }
            row.Cells["ColumnTitle"].Value = _richTextBox.Rtf;

            var imageSampleBuffer = dataRowView["ImageSample"];
            if (imageSampleBuffer != DBNull.Value)
                if ((imageSampleBuffer as byte[]).Length > 0)
                {
                    Image imageSample = GetImageFromBinary((byte[])imageSampleBuffer);
                    row.Cells["imageSample"].Value = ChangeImageOpacity(imageSample, 0.8f);
                }
            if (dataGridView.Columns["AppImage"].Visible)
            {
                var bitmap = ApplicationIcon(dataRowView["appPath"].ToString(), false);
                if (bitmap != null)
                    row.Cells["AppImage"].Value = bitmap;
            }
            dataGridView.Columns["AppPath"].Visible = false;
            UpdateTableGridRowBackColor(row);
        }

        private void PreviousMatchListMenuItem_Click(object sender, EventArgs e)
        {
            GotoSearchMatchInList(false);
        }

        private bool ProcessEnterKeyDown(bool isControlPressed, bool isShiftPressed)
        {
            PasteMethod pasteMethod;
            if (isControlPressed && !isShiftPressed)
                pasteMethod = PasteMethod.Text;
            else if (isControlPressed && isShiftPressed)
                pasteMethod = PasteMethod.Line;
            else
            {
                //if (!pasteENTERToolStripMenuItem.Enabled)
                if (richTextBox.Focused && EditMode)
                    return true;
                pasteMethod = PasteMethod.Standard;
            }
            SendPasteOfSelectedTextOrSelectedClips(pasteMethod);
            return false;
        }

        private void ReadSearchString()
        {
            searchString = comboBoxSearchString.Text;
        }

        private string RegexpPattern()
        {
            string textPattern = "";
            if (!String.IsNullOrEmpty(searchString))
                textPattern = RegexpPatternFromTextFilter();
            else
            {
                string filterValue = TypeFilter.SelectedValue as string;
                if (filterValue.Contains("text") && filterValue != "text")
                    textPattern = TextPatterns[filterValue.Substring("text_".Length)];
            }
            return textPattern;
        }

        private string RegexpPatternFromTextFilter()
        {
            string result = searchString;
            string[] array;
            if (settings.SearchWordsIndependently)
                array = result.Split(new[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            else
                array = new string[1] { result };
            result = "";
            foreach (var word in array)
            {
                if (result != "")
                    result += "|";
                result += "(" + Regex.Escape(word) + ")";
            }
            if (settings.SearchWildcards)
                result = result.Replace("%", ".*?");
            if (!String.IsNullOrWhiteSpace(result))
                result = "(" + result + ")";
            return result;
        }

        private void RegisterClipIdChange(int oldID, int newID)
        {
            if (lastPastedClips.ContainsKey(oldID))
            {
                var value = lastPastedClips[oldID];
                lastPastedClips.Remove(oldID);
                lastPastedClips[newID] = value;
            }
        }

        private async void ReloadList(bool forceRowLoad = false, int currentClipId = 0, bool keepTextSelectionIfIDChanged = false, List<int> selectedClipIDs = null, bool waitFinish = false)
        {
            //if (globalDataAdapter == null)
            //    return;
            if (!(this.Visible && this.ContainsFocus))
                sortField = "Id";
            if (EditMode)
                SaveClipText();
            string TypeFilterSelectedValue = TypeFilter.SelectedValue as string;
            if (currentClipId == 0 && clipBindingSource.Current != null)
            {
                currentClipId = (int)(clipBindingSource.Current as DataRowView).Row["Id"];
                if (dataGridView.SelectedRows.Count > 1 && selectedClipIDs == null)
                {
                    selectedClipIDs = new List<int>();
                    foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
                    {
                        if (selectedRow == null)
                            continue;
                        DataRowView dataRow = (DataRowView)selectedRow.DataBoundItem;
                        selectedClipIDs.Insert(0, (int)dataRow["Id"]);
                    }
                }
            }
            string MarkFilterSelectedValue = MarkFilter.SelectedValue as string;
            DateTime monthCalendar1SelectionStart = monthCalendar1.SelectionStart;
            DateTime monthCalendar1SelectionEnd = monthCalendar1.SelectionEnd;

            Task<DataTable> reloadListTask = Task.Run(() => ReloadListAsync(TypeFilterSelectedValue, MarkFilterSelectedValue, monthCalendar1SelectionStart, monthCalendar1SelectionEnd));
            lastReloadListTask = reloadListTask;
            if (waitFinish)
                reloadListTask.Wait();
            DataTable table = await reloadListTask;
            if (true && lastReloadListTask != null && lastReloadListTask != reloadListTask)
            {
                return;
            }
            lastReloadListTime = DateTime.Now;

            clipBindingSource.DataSource = table;
            stripLabelPosition.Spring = false;
            stripLabelPosition.Width = 50;
            stripLabelFiltered.Visible = filterOn;
            if (filterOn)
                stripLabelFiltered.Text = String.Format(Properties.Resources.FilteredStatusText, table.Rows.Count);
            stripLabelPosition.Spring = true;
            //PrepareTableGrid(); // Long
            if (filterOn)
            {
                toolStripButtonClearFilter.Enabled = true;
                //toolStripButtonClearFilter.Checked = true; // Back color wil not change
                toolStripButtonClearFilter.BackColor = Color.GreenYellow;
            }
            else
            {
                toolStripButtonClearFilter.Enabled = false;
                toolStripButtonClearFilter.BackColor = DefaultBackColor;
            }
            if (LastId == 0)
            {
                GotoLastRow();
                DataRowView lastRow = (DataRowView)clipBindingSource.Current;
                if (lastRow == null)
                {
                    LastId = 0;
                }
                else
                {
                    LastId = (int)lastRow["Id"];
                }
            }
            else
            {
                allowRowLoad = false;
                dataGridView.ClearSelection();
                allowRowLoad = true;
                RestoreSelectedCurrentClip(forceRowLoad, currentClipId, false, keepTextSelectionIfIDChanged);
                if (selectedClipIDs != null)
                {
                    allowProcessDataGridSelectionChanged = false;
                    foreach (int selectedID in selectedClipIDs)
                    {
                        SelectRowByID(selectedID);
                    }
                    allowProcessDataGridSelectionChanged = true;
                }
            }
            allowRowLoad = true;
        }

        private void ResetIsMainProperty()
        {
            SetProp(this.Handle, IsMainPropName, new IntPtr(1));
        }

        private void RestoreSelectedCurrentClip(bool forceRowLoad = false, int currentClipId = -1, bool clearSelection = true, bool keepTextSelectionIfIDChanged = false)
        {
            if (false
                //|| AutoGotoLastRow
                || currentClipId <= 0)
            {
                UpdateSelectedClipsHistory();
                GotoLastRow();
            }
            else if (currentClipId > 0)
            {
                int newPosition = clipBindingSource.Find("Id", currentClipId);
                allowRowLoad = false;
                if (newPosition == -1)
                {
                    UpdateSelectedClipsHistory();
                    GotoLastRow();
                }
                else
                {
                    // Calls SelectionChanged in DataGridView. Resets selectedRows!
                    clipBindingSource.Position = newPosition;
                }
                allowRowLoad = true;
                SelectCurrentRow(forceRowLoad || newPosition == -1, true, clearSelection, keepTextSelectionIfIDChanged);
            }
        }

        private void RestoreWindowIfMinimized(int newX = -12345, int newY = -12345, bool safeOpen = false)
        {
            this.FormBorderStyle = FormBorderStyle.Sizable;
            if (!allowVisible)
            {
                // executed only in first call of process life
                allowVisible = true;
                Show();
            }
            UpdateWindowTitle(true);
            if (newX == -12345)
            {
                if (this.Left > maxWindowCoordForHiddenState)
                    newX = this.Left;
                else
                    newX = this.RestoreBounds.X;
            }
            if (newY == -12345)
            {
                if (this.Top > maxWindowCoordForHiddenState)
                    newY = this.Top;
                else
                    newY = this.RestoreBounds.Y;
            }
            if (!safeOpen && settings.FastWindowOpen)
            {
                if (newY <= maxWindowCoordForHiddenState)
                    newY = factualTop;
                if (newX > maxWindowCoordForHiddenState)
                    MoveWindow(this.Handle, newX, newY, this.Width, this.Height, true);
            }
            else
            {
                this.Left = newX;
                this.Top = newY;
            }
            if (this.WindowState == FormWindowState.Minimized)
                this.WindowState = FormWindowState.Normal; // Window can be minimized by "Minimize All" command
            this.Activate(); // Without it window can be shown and be not focused
        }

        private void RichText_Click(object sender, EventArgs e)
        {
            OpenLinkIfAltPressed(richTextBox, e, TextLinkMatches);
            if (MaxTextViewSize >= richTextBox.SelectionStart && TextWasCut)
                AfterRowLoad(true);
        }

        private void richTextBox_Enter(object sender, EventArgs e)
        {
            if (true
                && clip != null
                && clip.Type == "file"
                && richTextBox.SelectionLength == 0
                && richTextBox.SelectionStart == 0)
            {
                var match = Regex.Match(richTextBox.Text, @"([^\\/:*?""<>|\r\n]+)[$<\r\n]", RegexOptions.Singleline);
                if (match != null)
                {
                    SetRichTextboxSelection(match.Groups[1].Index, match.Groups[1].Length);
                }
            }
        }

        private void richTextBox_SelectionChanged(object sender = null, EventArgs e = null)
        {
            if (!allowTextPositionChangeUpdate || htmlMode)
                return;
            if (!EditMode && richTextBox.SelectionStart + richTextBox.SelectionLength > clipRichTextLength)
            {
                richTextBox.Select(richTextBox.SelectionStart, clipRichTextLength - richTextBox.SelectionStart);
                return;
            }
            selectionStart = richTextBox.SelectionStart;
            if (selectionStart > richTextBox.Text.Length)
                selectionStart = richTextBox.Text.Length;
            int line = richTextBox.GetLineFromCharIndex(selectionStart);
            int lineStart = richTextBox.GetFirstCharIndexFromLine(line);
            string strLine = richTextBox.Text.Substring(lineStart, selectionStart - lineStart);
            char tab = '\u0009'; // TAB
            var TabSpace = new String(' ', tabLength);
            strLine = strLine.Replace(tab.ToString(), TabSpace);
            int column = strLine.Length + 1;
            line++;
            //if (richTextBox.Text.Length - MultiLangEndMarker().Length < (int)RowReader["chars"])
            //    selectionStart += line - 1; // to take into account /r/n vs /n
            selectionLength = richTextBox.SelectionLength;
            UpdateClipContentPositionIndicator(line, column);
        }

        private void rtfMenuItemOpenLink_Click(object sender, EventArgs e)
        {
            OpenLinkFromTextBox(TextLinkMatches, richTextBox.SelectionStart);
        }

        private void saveAsFileMenuItem_Click(object sender, EventArgs e)
        {
            string tempFile = GetClipTempFile(out string fileEditor);
            if (String.IsNullOrEmpty(tempFile))
                return;
            string extension = Path.GetExtension(tempFile);
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            string fileName = RemoveInvalidCharsFromFileName(clip.Title.Substring(0,50) + " " + clip.Created);
            saveFileDialog.FileName = fileName;
            saveFileDialog.CheckFileExists = false;
            saveFileDialog.Filter = extension + "| *" + extension + "|All|*.*";
            if (saveFileDialog.ShowDialog(this) != DialogResult.OK)
                return;
            File.Copy(tempFile, saveFileDialog.FileName);
        }

        private void SaveClipText()
        {
            int byteSize = 0;
            int chars = 0;
            string newText = richTextBox.Text;
            CalculateByteAndCharSizeOfClip("", "", newText, ref chars, ref byteSize);
            //string sql = "Update Clips set Title = @Title, Text = @Text, Size = @Size, Chars = @Chars where Id = @Id";
            //SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
            //command.Parameters.AddWithValue("@Id", clip.Id);
            //command.Parameters.AddWithValue("@Text", newText);
            //command.Parameters.AddWithValue("@Size", byteSize);
            //command.Parameters.AddWithValue("@chars", chars);
            string newTitle = "";
            if (clip.Title == TextClipTitle(clip.Text))
                newTitle = TextClipTitle(newText);
            else
                newTitle = clip.Title;
            //command.Parameters.AddWithValue("@Title", newTitle);
            //command.ExecuteNonQuery();

            var clip_get = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Id == clip.Id);
            if (clip_get != null)
            {
                clip_get.Text = newText;
                clip_get.Size = byteSize;
                clip_get.Chars = chars;
                clip_get.Title = newTitle;
                liteDb.GetCollection<Clip>("clips").Update(clip_get);
            }
        }

        private void SaveFilterInLastUsedList()
        {
            List<string> lastFilterValues = settings.LastFilterValues;
            if (!String.IsNullOrEmpty(searchString) && !lastFilterValues.Contains(searchString))
            {
                lastFilterValues.Insert(0, searchString);
                while (lastFilterValues.Count > 20)
                {
                    lastFilterValues.RemoveAt(lastFilterValues.Count - 1);
                }
                FillFilterItems();
            }
        }

        private void searchAllFieldsMenuItem_Click(object sender, EventArgs e)
        {
            settings.SearchAllFields = !settings.SearchAllFields;
            UpdateControlsStates();
            SearchStringApply();
        }

        private void SearchString_KeyDown(object sender, KeyEventArgs e)
        {
            PassKeyToGrid(true, e);
        }

        private void SearchString_KeyPress(object sender, KeyPressEventArgs e)
        {
            // http://csharpcoding.org/tag/keypress/ Workaroud strange beeping
            if (e.KeyChar == (char)Keys.Enter || e.KeyChar == (char)Keys.Escape)
                e.Handled = true;
        }

        private void SearchString_KeyUp(object sender, KeyEventArgs e)
        {
            PassKeyToGrid(false, e);
        }

        private void SearchString_TextChanged(object sender, EventArgs e)
        {
            if (AllowFilterProcessing || !settings.FilterListBySearchString)
            {
                timerApplySearchString.Stop();
                timerApplySearchString.Start();
            }
        }

        private void SearchStringApply()
        {
            ReadSearchString();
            searchMatchedIDs.Clear();
            if (settings.FilterListBySearchString)
            {
                ReloadList(true);
            }
            else

            {
                UpdateSearchMatchedIDs();
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    row.Cells["ColumnTitle"].Value = null;
                }
                if (settings.AutoSelectMatchedClip)
                {
                    GotoSearchMatchInList(true, true);
                }
                else
                {
                    SelectCurrentRow(true);
                }
            }
        }

        private void SelectCurrentRow(bool forceRowLoad = false, bool keepTextSelection = true, bool clearSelection = true, bool keepTextSelectionIfIDChanged = false)
        {
            if (clearSelection)
            {
                allowRowLoad = false;
                dataGridView.ClearSelection();
                allowRowLoad = true;
            }
            if (dataGridView.CurrentRow == null)
            {
                GotoLastRow();
                return;
            }
            allowRowLoad = false;
            dataGridView.Rows[dataGridView.CurrentRow.Index].Selected = true;
            allowRowLoad = true;
            LoadClipIfChangedID(forceRowLoad, keepTextSelection, keepTextSelectionIfIDChanged);
        }

        private void SelectNextMatchInClipText(bool fromCurrentSelection = true)
        {
            if (htmlMode)
                SelectNextMatchInWebBrowser(1, fromCurrentSelection);
            else
            {
                RichTextBox control = richTextBox;
                if (FilterMatches == null)
                    return;
                foreach (Match match in FilterMatches)
                {
                    if (false
                        || control.SelectionStart < match.Index
                        || (true
                            && control.SelectionLength == 0
                            && match.Index == 0
                        ))
                    {
                        allowTextPositionChangeUpdate = false;
                        control.SelectionStart = match.Groups[1].Index;
                        control.SelectionLength = match.Groups[1].Length;
                        control.HideSelection = false;
                        allowTextPositionChangeUpdate = true;
                        OnClipContentSelectionChange();
                        break;
                    }
                }
            }
        }

        private void SelectNextMatchInWebBrowser(int direction, bool fromCurrentSelection = true)
        {
            IHTMLDocument2 htmlDoc = (IHTMLDocument2)htmlTextBox.Document.DomDocument;
            IHTMLBodyElement body = htmlDoc.body as IHTMLBodyElement;
            IHTMLTxtRange currentRange = null;
            if (fromCurrentSelection)
                currentRange = GetHtmlCurrentTextRangeOrAllDocument();
            IHTMLTxtRange nearestMatch = null;
            int searchFlags = 0;
            if (settings.SearchCaseSensitive)
                searchFlags = 4;
            string[] array;
            if (settings.SearchWordsIndependently)
                array = searchString.Split(new[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            else
                array = new string[1] { searchString };
            foreach (var word in array)
            {
                IHTMLTxtRange range = body.createTextRange();
                if (currentRange != null)
                {
                    if (direction > 0)
                        range.setEndPoint("StartToEnd", currentRange);
                    else
                        range.setEndPoint("EndToStart", currentRange);
                }
                if (range.findText(word, direction, searchFlags))
                {
                    if (false
                        || nearestMatch == null
                        || (true
                            && direction > 0
                            && (false
                                ||
                                (nearestMatch as IHTMLTextRangeMetrics).boundingTop > (range as IHTMLTextRangeMetrics).boundingTop
                                || (true
                                    && (nearestMatch as IHTMLTextRangeMetrics).boundingTop == (range as IHTMLTextRangeMetrics).boundingTop
                                    && (nearestMatch as IHTMLTextRangeMetrics).boundingLeft > (range as IHTMLTextRangeMetrics).boundingLeft)))
                        || (true
                            && direction < 0
                            && (false
                                || (nearestMatch as IHTMLTextRangeMetrics).boundingTop < (range as IHTMLTextRangeMetrics).boundingTop
                                || (true
                                    && (nearestMatch as IHTMLTextRangeMetrics).boundingTop == (range as IHTMLTextRangeMetrics).boundingTop
                                    && (nearestMatch as IHTMLTextRangeMetrics).boundingLeft < (range as IHTMLTextRangeMetrics).boundingLeft))))
                    {
                        nearestMatch = range;
                    }
                }
            }
            if (nearestMatch != null)
            {
                nearestMatch.scrollIntoView();
                nearestMatch.@select();
            }
        }

        private void SelectRowByID(int IDToSelect)
        {
            int newIndex = clipBindingSource.Find("Id", IDToSelect);
            if (newIndex >= 0)
            {
                dataGridView.Rows[newIndex].Selected = false;
                dataGridView.Rows[newIndex].Selected = true;
            }
        }

        private IHTMLTxtRange SelectTextRangeInWebBrowser(int NewSelectionStart, int NewSelectionLength)
        {
            IHTMLDocument2 htmlDoc = htmlTextBox.Document.DomDocument as IHTMLDocument2;
            IHTMLBodyElement body = htmlDoc.body as IHTMLBodyElement;
            IHTMLTxtRange range = body.createTextRange();
            range.moveStart("character", NewSelectionStart);
            range.collapse();
            range.moveEnd("character", NewSelectionLength);
            range.@select();
            return range;
        }

        private void sendKey(IntPtr hwnd, Keys keyCode, bool extended = false, bool down = true, bool up = true)
        {
            // http://stackoverflow.com/questions/10280000/how-to-create-lparam-of-sendmessage-wm-keydown
            const int WM_KEYDOWN = 0x0100;
            const int WM_KEYUP = 0x0101;
            uint scanCode = MapVirtualKey((uint)keyCode, 0);
            uint lParam = 0x00000001 | (scanCode << 16);
            if (extended)
            {
                lParam |= 0x01000000;
            }
            if (down)
            {
                PostMessage(hwnd, WM_KEYDOWN, (int)keyCode, (int)lParam);
            }
            lParam |= 0xC0000000; // set previous key and transition states (bits 30 and 31)
            if (up)
            {
                PostMessage(hwnd, WM_KEYUP, (int)keyCode, (int)lParam);
            }
        }

        // Return - bool - true if failed
        private bool SendPaste(PasteMethod pasteMethod = PasteMethod.Standard)
        {
    
            _ = GetWindowThreadProcessId(lastActiveParentWindow, out int targetProcessId);
            bool needElevation = targetProcessId != 0 && !UacHelper.IsProcessAccessible(targetProcessId);
            var curproc = Process.GetCurrentProcess();
            if (needElevation)
            {
                string ElevatedMutexName = "ClippyElevatedMutex" + curproc.Id;
                Mutex ElevatedMutex = null;
                try
                {
                    ElevatedMutex = Mutex.OpenExisting(ElevatedMutexName);
                }
                catch
                {
                    string exePath = curproc.MainModule.FileName;
                    ProcessStartInfo startInfo = new(exePath, "/elevated " + curproc.Id);
                    startInfo.Verb = "runas";
                    try
                    {
                        Process.Start(startInfo);
                    }
                    catch
                    {
                        ShowElevationFail();
                        return true;
                    }
                }
                int maxWait = 2000;
                Stopwatch stopWatch = new();
                stopWatch.Start();
                while (stopWatch.ElapsedMilliseconds < maxWait)
                {
                    try
                    {
                        ElevatedMutex = Mutex.OpenExisting(ElevatedMutexName);
                        break;
                    }
                    catch
                    {
                    }
                    Thread.Sleep(5);
                }
                if (ElevatedMutex == null)
                {
                    ShowElevationFail();
                    return true;
                }
            }
            ActivateAndCheckTargetWindow();
            bool targetIsCurrentProcess = DoActiveWindowBelongsToCurrentProcess(IntPtr.Zero);
            if (targetIsCurrentProcess)
                return true;
            if (pasteMethod == PasteMethod.SendCharsFast || pasteMethod == PasteMethod.SendCharsSlow)
            {
                if (!IsTextType())
                    return true;
                if (!needElevation)
                    Paster.SendChars(this, pasteMethod == PasteMethod.SendCharsSlow);
                else
                {
                    EventWaitHandle sendCharsEvent = Paster.GetSendCharsEventWaiter(0, pasteMethod == PasteMethod.SendCharsSlow);
                    sendCharsEvent.Set();
                }
            }
            else
            {
                if (settings.DontSendPaste)
                    return false;
                if (!needElevation)
                {
                    Paster.SendPaste(this);
                }
                else
                {
                    EventWaitHandle pasteEvent = Paster.GetPasteEventWaiter();
                    pasteEvent.Set();
                }
                lastPasteMoment = DateTime.Now;
            }
            return false;
        }

        // Does not respect MoveCopiedClipToTop
        private string SendPasteClipExpress(DataGridViewRow currentViewRow = null, PasteMethod pasteMethod = PasteMethod.Standard, bool pasteDelimiter = false, bool updateDB = false, string DelimiterForTextJoin = null)
        {
            if (currentViewRow == null)
                currentViewRow = dataGridView.CurrentRow;
            if (currentViewRow == null)
                return "";
            var dataRow = (DataRowView)currentViewRow.DataBoundItem;

            //var rowReader = (BsonDataReader)liteDb.GetCollection<Clip>("clips").Query()
            //    .Where(x => x.Id == (int)dataRow["id"])
            //    .ExecuteReader();

            int rowId = (int)dataRow["id"];

            Clip rowReader;

            try
            {
                rowReader = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Id==rowId);
            }
            catch (Exception)
            {

                return "";
            }

            //var rowReader = getRowReader((int)dataRow["id"]);
            string type = rowReader.Type;
            if (pasteMethod == PasteMethod.Null)
            {
                string textToPaste = rowReader.Text;
                if (type == "img")
                {
                    string fileEditor = "";
                    textToPaste = GetClipTempFile(out fileEditor, rowReader);
                }
                if (pasteDelimiter)
                {
                    if (DelimiterForTextJoin == null)
                        DelimiterForTextJoin = settings.DelimiterForTextJoin;
                    textToPaste = DelimiterForTextJoin.Replace("\\n", Environment.NewLine) + textToPaste;
                }
                return textToPaste;
            }
            if (pasteDelimiter && pasteMethod == PasteMethod.Standard)
            {
                int multipasteDelay = 50;
                Thread.Sleep(multipasteDelay);

                SetTextInClipboard(Environment.NewLine + Environment.NewLine, false);
                SendPaste(pasteMethod);
                Thread.Sleep(multipasteDelay);
            }
            CopyClipToClipboard(rowReader, pasteMethod != PasteMethod.Standard && pasteMethod != PasteMethod.File, false);
            if (SendPaste(pasteMethod))
                return "";

            if (updateDB)
                SetRowMark("Used", true, false, true);

            if (currentViewRow.DataBoundItem != null)
            {
                ((DataRowView)currentViewRow.DataBoundItem).Row["Used"] = true;
                UpdateTableGridRowBackColor(currentViewRow);
            }
            return "";
        }

        private void SendPasteOfSelectedTextOrSelectedClips(PasteMethod pasteMethod = PasteMethod.Standard)
        {
            string agregateTextToPaste = "";
            string selectedText = "";
            PasteMethod itemPasteMethod;
            if (pasteMethod == PasteMethod.File)
            {
                DataObject dto = new DataObject();
                string clipText = SetClipFilesInDataObject(dto);
                SetTextInClipboardDataObject(dto, clipText);
                SetClipboardDataObject(dto, false);
                SendPaste();
            }
            else
            {
                if (pasteMethod == PasteMethod.Standard)
                    itemPasteMethod = pasteMethod;
                else
                    itemPasteMethod = PasteMethod.Null;
                agregateTextToPaste = GetSelectedTextOfClips(ref selectedText, itemPasteMethod);
                if (itemPasteMethod == PasteMethod.Null && !String.IsNullOrEmpty(agregateTextToPaste))
                {
                    if (pasteMethod == PasteMethod.Line)
                    {
                        agregateTextToPaste = ConvertTextToLine(agregateTextToPaste);
                    }
                    SetTextInClipboard(agregateTextToPaste, false);
                    SendPaste(pasteMethod);
                }
            }

            if (String.IsNullOrEmpty(selectedText))
            {
                SetRowMark("Used", true, true, true);
            }
            if (true
                && settings.MoveCopiedClipToTop
                && String.IsNullOrEmpty(selectedText)
                )
            {
                MoveSelectedRows(0);
                //CaptureClipboardData();
            }
            else if (true
                     && pasteMethod == PasteMethod.Text
                     && !String.IsNullOrEmpty(selectedText))
            {
                // With multipaste works incorrect
                CaptureClipboardData();
                if (settings.MoveCopiedClipToTop)
                    MoveSelectedRows(0);
            }
        }

        private void SetClipboardDataObject(IDataObject dto, bool allowSelfCapture = true)
        {
            // If not doing this, WM_CLIPBOARDUPDATE event will be raised 2 times (why?) if "copy"=true
            RemoveClipboardFormatListener(this.Handle);
            //bool success = false;
            try
            {
                Clipboard.SetDataObject(dto, true, 10, 20);
                // Very important to set second parameter to TRUE to give immidiate access to buffer to other processes!
                //success = true;
            }
            catch
            {
                ClipboardOwner clipboardOwner = GetClipboardOwnerLockerInfo(true);
                Debug.WriteLine(String.Format(Properties.Resources.FailedToWriteClipboard, clipboardOwner.windowTitle, clipboardOwner.application));
            }

            ConnectClipboard();
            if (allowSelfCapture)
                CaptureClipboardData();
        }

        private string SetClipFilesInDataObject(DataObject dto, int maxRowsDrag = 100)
        {
            string textList = "";
            StringCollection fileNameCollection = new StringCollection();
            foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
            {
                DataRowView dataRowView = (DataRowView)selectedRow.DataBoundItem;

                //var RowReader = (BsonDataReader)liteDb.GetCollection<Clip>("clips").Query()
                //                    .Where(x => x.Id == (int)dataRowView["id"])
                //                    .ExecuteReader();

                int rowId = (int)dataRowView["id"];
                var localClip = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Id == rowId);
        

                //BsonDataReader RowReader = getRowReader((int)dataRowView["id"]);
                string clipText;
                DataObject clipDto = ClipDataObject(localClip, false, out clipText);
                if (localClip.Type != "file")
                {
                    string junkVar;
                    string filename = GetClipTempFile(out junkVar, localClip);
                    fileNameCollection.Add(filename);
                    textList += filename;
                }
                else
                {
                    foreach (string filename in clipDto.GetFileDropList())
                    {
                        fileNameCollection.Add(filename);
                        textList += filename;
                    }
                }
                maxRowsDrag--;
                if (maxRowsDrag == 0)
                    break;
            }
            dto.SetFileDropList(fileNameCollection);
            return textList;
        }

        private void SetRichTextboxSelection(int NewSelectionStart, int NewSelectionLength, bool preventHardScroll = false)
        {
            richTextBox.SelectionStart = NewSelectionStart;
            richTextBox.SelectionLength = NewSelectionLength;
            if (preventHardScroll)
                richTextBox.HideSelection = true; // slow // Exeption in ScrollToCaret can be thrown without this
            try
            {
                richTextBox.ScrollToCaret();
            }
            catch
            {
                // Happens when click in not full loaded richTextBox
            }
            if (preventHardScroll)
                richTextBox.HideSelection = false; // slow
        }

        private void SetRowMark(string fieldName, bool newValue = true, bool allSelected = false, bool addToLastPasted = false)
        {
            //string sql = "Update Clips set " + fieldName + "=@Value where Id IN(null";
            //SQLiteCommand command = new SQLiteCommand("", m_dbConnection);
            
            List<DataGridViewRow> selectedRows = new List<DataGridViewRow>();
            if (allSelected)
                foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
                    selectedRows.Add(selectedRow);
            else
                selectedRows.Add(dataGridView.CurrentRow);
            int counter = 0;
            ReadSearchString();

            List<int> selectedIDs = new List<int>();
            foreach (DataGridViewRow selectedRow in selectedRows)
            {
                if (selectedRow == null)
                    continue;
                DataRowView dataRow = (DataRowView)selectedRow.DataBoundItem;
                //string parameterName = "@Id" + counter;
                //sql += "," + parameterName;
                //command.Parameters.Add(parameterName, DbType.Int32).Value = dataRow["Id"];

                selectedIDs.Insert(0, (int)dataRow["Id"]);
                counter++;
                
                dataRow[fieldName] = newValue;
                
                int rowId = (int)dataRow["Id"];
                var clip = liteDb.GetCollection<Clip>("clips").FindOne(x => x.Id == rowId);

                //tohle ale nefunguje:

                if (fieldName.ToLower() == "favorite")
                    clip.Favorite = newValue;

                if (fieldName.ToLower() == "used")
                    clip.Used = newValue;

                var col = liteDb.GetCollection<Clip>("clips").Update(clip);


                UpdateTableGridRowBackColor(selectedRow);
                if (addToLastPasted)
                {
                    lastPastedClips[(int)dataRow["Id"]] = DateTime.Now;
                    if (lastPastedClips.Count > 100)
                        lastPastedClips.Remove(lastPastedClips.Aggregate((l, r) => l.Value < r.Value ? l : r).Key);
                }
            }
            //sql += ")";
            //command.CommandText = sql;
            //command.Parameters.AddWithValue("@Value", newValue);
            //command.ExecuteNonQuery();

            if (allSelected)
                ReloadList(true);
            else
            {
                LoadRowReader();
                UpdateClipButtons();
            }
        }

        private void SetTextInClipboard(string text, bool allowSelfCapture = true)
        {
            DataObject dto = new DataObject();
            SetTextInClipboardDataObject(dto, text);
            SetClipboardDataObject(dto, allowSelfCapture);
        }

        void propertyGrid_PropertyValueChanged(object sender, PropertyValueChangedEventArgs e)
        {
            
            var set = liteDb.GetCollection<Settings>("settings").Update(settings);

            if (e.ChangedItem.Label.Equals("DatabasePath"))
            {
                RegistryKey key = Registry.CurrentUser.CreateSubKey("SOFTWARE\\Clippy", true);
                key.SetValue("dbPath", e.ChangedItem.Value);
            }

            if (e.ChangedItem.Label.Equals("Autostart"))
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);
                if ((bool)e.ChangedItem.Value == true)
                {
                    //key.SetValue("Clippy", "\"" + Application.ExecutablePath + "\"" + " -tray");
                    key.SetValue("Clippy", "\"" + Application.ExecutablePath + "\"");
                }
                if ((bool)e.ChangedItem.Value == false)
                {
                    try
                    {
                        key.DeleteValue("Clippy");
                    }
                    catch{}
                }
            }

        }

        private void ShowElevationFail()
        {
            MessageBox.Show(this, Properties.Resources.CantPasteInElevatedWindow, Application.ProductName);
        }

        private void ShowForPaste(bool onlyFavorites = false, bool clearFiltersAndGoToTop = false, bool safeOpen = false)
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form.Modal && form.Visible && form.CanFocus)
                {
                    form.Activate();
                    return;
                }
            }

            if (onlyFavorites)
                showOnlyFavoriteToolStripMenuItem_Click();

            if (clearFiltersAndGoToTop)
                ClearFilter(-1, onlyFavorites);

            int newX = -12345;
            int newY = -12345;
            if (this.Visible && this.ContainsFocus)
            {
                RestoreWindowIfMinimized(newX, newY, safeOpen);
                return;
            }
            // https://www.codeproject.com/Articles/34520/Getting-Caret-Position-Inside-Any-Application
            // http://stackoverflow.com/questions/31055249/is-it-possible-to-get-caret-position-in-word-to-update-faster
            UpdateLastActiveParentWindow(IntPtr.Zero); // sometimes lastActiveParentWindow is not equal GetForegroundWindow()
            IntPtr hWindow = lastActiveParentWindow;
            if (hWindow != IntPtr.Zero)
            {
                Point caretPoint;
                GUITHREADINFO guiInfo = GetGuiInfo(hWindow, out caretPoint);
                if (caretPoint.Y > 0 && settings.RestoreCaretPositionOnFocusReturn)
                {
                    lastChildWindow = guiInfo.hwndFocus;
                    GetWindowRect(lastChildWindow, out lastChildWindowRect);
                    lastCaretPoint = new Point(guiInfo.rcCaret.left, guiInfo.rcCaret.top);
                }
                if (settings.WindowAutoPosition)
                {
                    RECT activeRect;
                    if (caretPoint.Y > 0)
                    {
                        activeRect = guiInfo.rcCaret;
                        Screen screen = Screen.FromPoint(caretPoint);
                        newX = Math.Max(screen.WorkingArea.Left, Math.Min(activeRect.right + caretPoint.X, screen.WorkingArea.Width + screen.WorkingArea.Left - this.Width));
                        newY = Math.Max(screen.WorkingArea.Top, Math.Min(activeRect.bottom + caretPoint.Y + 1, screen.WorkingArea.Height + screen.WorkingArea.Top - this.Height));
                    }
                    else
                    {
                        IntPtr baseWindow;
                        if (guiInfo.hwndFocus != IntPtr.Zero)
                            baseWindow = guiInfo.hwndFocus;
                        else
                            baseWindow = hWindow;
                        ClientToScreen(baseWindow, out caretPoint);
                        GetWindowRect(baseWindow, out activeRect);
                        Screen screen = Screen.FromPoint(caretPoint);
                        newX = Math.Max(screen.WorkingArea.Left,
                            Math.Min((activeRect.right - activeRect.left - this.Width) / 2 + caretPoint.X, screen.WorkingArea.Width + screen.WorkingArea.Left - this.Width));
                        newY = Math.Max(screen.WorkingArea.Top,
                            Math.Min((activeRect.bottom - activeRect.top - this.Height) / 2 + caretPoint.Y, screen.WorkingArea.Height + screen.WorkingArea.Top - this.Height));
                    }
                }
            }

            RestoreWindowIfMinimized(newX, newY, safeOpen);
            if (!settings.FastWindowOpen || safeOpen)
            {
                this.Activate();
                this.Show();
            }
            SetForegroundWindow(this.Handle);
        }

        private void SwitchMonitoringClipboard(bool showToolTip = false)
        {
            settings.MonitoringClipboard = !settings.MonitoringClipboard;
            if (settings.MonitoringClipboard)
                ConnectClipboard();
            else
                RemoveClipboardFormatListener(this.Handle);
            UpdateControlsStates();
            UpdateWindowTitle();
            //UpdateNotifyIcon();
            if (showToolTip)
            {
                string text;
                if (settings.MonitoringClipboard)
                    text = Properties.Resources.MonitoringON;
                else
                    text = Properties.Resources.MonitoringOFF;
                notifyIcon.ShowBalloonTip(2000, Application.ProductName, text, ToolTipIcon.Info);
            }
        }

        private void textBoxUrl_Click(object sender, EventArgs e)
        {
            OpenLinkIfAltPressed(sender as RichTextBox, e, UrlLinkMatches);
        }

        private void timerApplySearchString_Tick(object sender = null, EventArgs e = null)
        {
            timerApplySearchString.Stop();
            SearchStringApply();
        }

        private void timerReconnect_Tick(object sender, EventArgs e)
        {
            ConnectClipboard();
        }

        private void TypeFilter_SelectedValueChanged(object sender, EventArgs e)
        {
            if (AllowFilterProcessing)
            {
                ReloadList(true);
            }
        }

        private void UpdateClipButtons()
        {
            //toolStripButtonMarkFavorite.Checked = BoolFieldValue("Favorite");

            if(BoolFieldValue("Favorite"))
            {
                toolStripMenuItemMarkFavorite.BackColor = Color.LightBlue;
            }
            else
            {
                toolStripMenuItemMarkFavorite.BackColor = Control.DefaultBackColor;
            }

            
            // dataGridView.CurrentRow could be null here!
        }

        private void UpdateClipContentPositionIndicator(int line = 1, int column = 1)
        {
            string newText;
            if (clip != null && clip.Type == "img")
            {
                //double zoomFactor = Math.Min((double) ImageControl.ClientSize.Width / ImageControl.Image.Width, (double) ImageControl.ClientSize.Height / ImageControl.Image.Height);
                double zoomFactor = ImageControl.ZoomFactor();
                newText = Properties.Resources.Zoom + ": " + zoomFactor.ToString("0.00");
            }
            else
            {
                newText = "" + selectionStart;
                newText += "(" + line + ":" + column + ")";
                if (selectionLength > 0)
                {
                    newText += "+" + selectionLength;
                }
            }
            stripLabelPosition.Text = newText;
        }

        private void UpdateColumnsSet()
        {
            dataGridView.Columns["appImage"].Visible = settings.ShowApplicationIconColumn;
            //dataGridView.Columns["VisualWeight"].Visible = settings.ShowVisualWeightColumn;
            dataGridView.Columns["ColumnCreated"].Visible = settings.ShowSecondaryColumns;
        }

        public void UpdateControlsStates()
        {
            searchAllFieldsMenuItem.Checked = settings.SearchAllFields;
            filterListBySearchStringMenuItem.Checked = settings.FilterListBySearchString;
            autoselectMatchedClipMenuItem.Checked = settings.AutoSelectMatchedClip;
            toolStripMenuItemSecondaryColumns.Checked = settings.ShowSecondaryColumns;
            //toolStripButtonSecondaryColumns.Checked = settings.ShowSecondaryColumns;
            toolStripMenuItemSearchCaseSensitive.Checked = settings.SearchCaseSensitive;
            toolStripMenuItemSearchWordsIndependently.Checked = settings.SearchWordsIndependently;
            toolStripMenuItemSearchWildcards.Checked = settings.SearchWildcards;
            ignoreBigTextsToolStripMenuItem.Checked = settings.SearchIgnoreBigTexts;
            //moveCopiedClipToTopToolStripButton.Checked = settings.MoveCopiedClipToTop;
            moveCopiedClipToTopToolStripMenuItem.Checked = settings.MoveCopiedClipToTop;
            toolStripButtonAutoSelectMatch.Checked = settings.AutoSelectMatch;
            trayMenuItemMonitoringClipboard.Checked = settings.MonitoringClipboard;
            toolStripMenuItemMonitoringClipboard.Checked = settings.MonitoringClipboard;
            //toolStripButtonTextFormatting.Checked = settings.ShowNativeTextFormatting;
            textFormattingToolStripMenuItem.Checked = settings.ShowNativeTextFormatting;
            //toolStripButtonMonospacedFont.Checked = settings.MonospacedFont;
            monospacedFontToolStripMenuItem.Checked = settings.MonospacedFont;
            wordWrapToolStripMenuItem.Checked = settings.WordWrap;
            //toolStripButtonWordWrap.Checked = settings.WordWrap;
            richTextBox.WordWrap = wordWrapToolStripMenuItem.Checked;
            showInTaskbarToolStripMenuItem.Checked = settings.ShowInTaskBar;
            this.ShowInTaskbar = settings.ShowInTaskBar;
            // After ShowInTaskbar change true->false all window properties are deleted. So we need to reset it.
            ResetIsMainProperty();
            editClipTextToolStripMenuItem.Checked = EditMode;
            //toolStripMenuItemEditClipText.Checked = EditMode;

            //var WordWrap = settings.WordWrap ? toolStripButtonWordWrap1.BackColor = Color.LightGray : toolStripButtonWordWrap1.BackColor = Control.DefaultBackColor;
            //var ShowNativeTextFormatting = settings.ShowNativeTextFormatting ? toolStripButtonTextFormatting1.BackColor = Color.LightGray : toolStripButtonTextFormatting1.BackColor = Control.DefaultBackColor;

            if (EditMode)
                toolStripMenuItemEditClipText1.BackColor = Color.LightBlue;
            else
                toolStripMenuItemEditClipText1.BackColor = Control.DefaultBackColor;


            if (settings.WordWrap)
                toolStripButtonWordWrap1.BackColor = Color.LightBlue;
            else
                toolStripButtonWordWrap1.BackColor = Control.DefaultBackColor;

            if (settings.ShowNativeTextFormatting)
                toolStripButtonTextFormatting1.BackColor = Color.LightBlue;
            else
                toolStripButtonTextFormatting1.BackColor = Control.DefaultBackColor;

            if (settings.MonospacedFont)
                toolStripButtonMonospacedFont1.BackColor = Color.LightBlue;
            else
                toolStripButtonMonospacedFont1.BackColor = Control.DefaultBackColor;

            if (settings.ShowSecondaryColumns)
                toolStripButtonSecondaryColumns1.BackColor = Color.LightBlue;
            else
                toolStripButtonSecondaryColumns1.BackColor = Control.DefaultBackColor;
            
            if (settings.MoveCopiedClipToTop)
                moveCopiedClipToTopToolStripButton1.BackColor = Color.LightBlue;
            else
                moveCopiedClipToTopToolStripButton1.BackColor = Control.DefaultBackColor;
        }

        private void UpdateIgnoreModulesInClipCapture()
        {
            ignoreModulesInClipCapture = new StringCollection();
            if (settings.IgnoreApplicationsClipCapture != null)
                foreach (var fullFilename in settings.IgnoreApplicationsClipCapture)
                {
                    ignoreModulesInClipCapture.Add(Path.GetFileNameWithoutExtension(fullFilename).ToLower());
                }
        }

        private void UpdateLastActiveParentWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
                hwnd = GetForegroundWindow();
            bool targetIsCurrentProcess = DoActiveWindowBelongsToCurrentProcess(hwnd);
            if (!targetIsCurrentProcess && lastActiveParentWindow != hwnd)
            {
                //preLastActiveParentWindow = lastActiveParentWindow; // In case with Explorer tray click it will not help
                lastActiveParentWindow = hwnd;
                lastChildWindow = IntPtr.Zero;
                lastChildWindowRect = new RECT();
                lastCaretPoint = new Point();
                UpdateWindowTitle();
            }
        }

        private void UpdateSearchMatchedIDs()
        {

            var col = liteDb.GetCollection<Clip>("clips");

            searchMatchedIDs.Clear();
            if (!String.IsNullOrEmpty(searchString))
            {
                var results = col.Query()
                .Where(x => x.Text.Contains(searchString) || x.Title.Contains(searchString) || x.Window.Contains(searchString) || x.Url.Contains(searchString))
                .Select(x => x.Id )
                .ToList();

                if (results.Count > 0)
                {
                    searchMatchedIDs.AddRange(results);
                }

                //SQLiteCommand command = new SQLiteCommand(m_dbConnection);
                //command.CommandText = "Select Id From Clips";
                //command.CommandText += " WHERE 1=1 " + SqlSearchFilter();
                //command.CommandText += " ORDER BY " + sortField + " desc";
                //if (settings.SearchCaseSensitive)
                //    command.CommandText = "PRAGMA case_sensitive_like = 1; " + command.CommandText;
                //else
                //    command.CommandText = "PRAGMA case_sensitive_like = 0; " + command.CommandText;
                //SQLiteDataReader reader = command.ExecuteReader();
                //while (reader.Read())
                //{
                //    searchMatchedIDs.Add((int)reader["id"]);
                //}
            }
        }

        private void UpdateSelectedClipsHistory()
        {
            if (clip != null)
            {
                int oldID = clip.Id;
                if (!selectedClipsBeforeFilterApply.Contains(oldID))
                    selectedClipsBeforeFilterApply.Add(oldID);
            }
        }

        private void UpdateTableGridRowBackColor(DataGridViewRow row = null)
        {
            if (row == null)
                row = dataGridView.CurrentRow;
            DataRowView dataRowView = (DataRowView)(row.DataBoundItem);
            bool fav = BoolFieldValue("Favorite", dataRowView);
            bool used;
            try
            {
                used = (bool)dataRowView.Row["Used"];
            }
            catch (Exception)
            {
                // Not yet fully read
                return;
            }

            if (fav)
                row.DefaultCellStyle.BackColor = favoriteColor;
            else if (used)
                row.DefaultCellStyle.BackColor = _usedColor;
            else
                row.DefaultCellStyle.BackColor = default;

            //foreach (DataGridViewCell cell in row.Cells)
            //{
            //    if (fav)
            //    {

            //        cell.Style.BackColor = favoriteColor;
            //        //toolStripButtonMarkFavorite.Checked = true;
            //    }
            //    else if (used)
            //        cell.Style.BackColor = _usedColor;
            //    else
            //    {
            //        cell.Style.BackColor = default(Color);
            //        //toolStripButtonMarkFavorite.Checked = false;
            //    }

            //}
        }
        
        private void UpdateWindowTitle(bool forced = false)
        {
            if ((this.Top <= maxWindowCoordForHiddenState || !this.Visible) && !forced)
                return;
            string targetTitle = "<" + Properties.Resources.NoActiveWindow + ">";

            targetTitle = GetWindowTitle(lastActiveParentWindow);
            int pid;
            _ = GetWindowThreadProcessId(lastActiveParentWindow, out pid);
            Process proc = Process.GetProcessById(pid);
            if (proc != null)
            {
                targetTitle += " [" + proc.ProcessName + "]";
            }

            Debug.WriteLine("Active window " + lastActiveParentWindow + " " + targetTitle);
            string newTitle = Application.ProductName + " " + Properties.Resources.VersionValue;
            if (!settings.MonitoringClipboard)
            {
                newTitle += " [" + Properties.Resources.NoCapture + "]";
            }
            this.Text = newTitle + " >> " + targetTitle;
            notifyIcon.Text = newTitle;
        }



        public struct removeClipsFilter
        {
            public int PID;
            public DateTime TimeEnd;
            public DateTime TimeStart;
        }

        public class ClipboardOwner
        {
            public string application = "";
            public string appPath = "";
            public bool isRemoteDesktop = false;
            public int processId = 0;
            public string windowTitle = "";
        }

        internal sealed class KeyboardLayout
        {
            //http://stackoverflow.com/questions/37291533/change-keyboard-layout-from-c-sharp-code-with-net-4-5-2
            private readonly uint hkl;

            private KeyboardLayout(CultureInfo cultureInfo)
            {
                string layoutName = cultureInfo.LCID.ToString("x8");

                var pwszKlid = new StringBuilder(layoutName);
                this.hkl = LoadKeyboardLayout(pwszKlid,
                    KeyboardLayoutFlags.KLF_ACTIVATE | KeyboardLayoutFlags.KLF_SUBSTITUTE_OK);
            }

            private KeyboardLayout(uint hkl)
            {
                this.hkl = hkl;
            }

            public uint Handle
            {
                get { return this.hkl; }
            }

            public static KeyboardLayout GetCurrent()
            {
                uint hkl = GetKeyboardLayout((uint)Environment.CurrentManagedThreadId);
                return new KeyboardLayout(hkl);
            }

            public static KeyboardLayout Load(CultureInfo culture)
            {
                return new KeyboardLayout(culture);
            }

            public void Activate()
            {
                ActivateKeyboardLayout(this.hkl, KeyboardLayoutFlags.KLF_SETFORPROCESS);
            }

           

            private static class KeyboardLayoutFlags
            {
                //https://msdn.microsoft.com/ru-ru/library/windows/desktop/ms646305(v=vs.85).aspx
                public const uint KLF_ACTIVATE = 0x00000001;

                public const uint KLF_SETFORPROCESS = 0x00000100;
                public const uint KLF_SUBSTITUTE_OK = 0x00000002;
            }
        }

        //private void MergeCellsInRow(DataGridView dataGridView1, DataGridViewRow row, int col1, int col2)
        //{
        //    Graphics g = dataGridView1.CreateGraphics();
        //    Pen p = new Pen(dataGridView1.GridColor);
        //    Rectangle r1 = dataGridView1.GetCellDisplayRectangle(col1, row.Index, true);
        //    //Rectangle r2 = dataGridView1.GetCellDisplayRectangle(col2, row.Index, true);
        private class ListItemNameText
        {
            public string Name { get; set; }
            public string Text { get; set; }
        }
        
        private class MyToolStripRenderer : ToolStripProfessionalRenderer
        {
            protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e)
            {
                if (e.Item is ToolStripButton btn && btn.Checked && !btn.Selected)
                {
                    Rectangle bounds1 = new(Point.Empty, e.Item.Size);
                    bounds1 = new(bounds1.Left, bounds1.Top, e.Item.Size.Width - 1, e.Item.Size.Height - 1);
                    Rectangle bounds2 = new(bounds1.Left + 1, bounds1.Top + 1, bounds1.Right - 1,
                        bounds1.Bottom - 1);
                    e.Graphics.DrawRectangle(new Pen(Color.DeepSkyBlue), bounds1);
                    Brush brush = new SolidBrush(Color.FromArgb(255, 180, 240, 240));
                    e.Graphics.FillRectangle(brush, bounds2);
                }
                else base.OnRenderButtonBackground(e);
            }
        }


    }

    public partial class Clip
    {
        [BsonId]
        public int _id { get; set; }
        public int Id { get; set; }
        public string Type { get; set; }
        public string Text { get; set; }
        public string Title { get; set; }
        public string RichText { get; set; }
        public string HtmlText { get; set; }
        public string Application { get; set; }
        public string AppPath { get; set; }
        public string Window { get; set; }
        public int Chars { get; set; }
        public int Size { get; set; }
        public DateTime Created { get; set; }
        public Byte[] Binary { get; set; }
        public Byte[] ImageSample { get; set; }
        public string Url { get; set; }
        public string Hash { get; set; }
        public bool Favorite { get; set; } = false;
        public bool Used { get; set; } = false;
        public bool Contain_time { get; set; } = false;
        public bool Contain_email { get; set; } = false;
        public bool Contain_number { get; set; } = false;
        public bool Contain_phone { get; set; } = false;
        public bool Contain_url { get; set; } = false;
        public bool Contain_url_image { get; set; } = false;
        public bool Contain_url_video { get; set; } = false;
        public bool Contain_filename { get; set; } = false;
    }

    public partial class Settings
    {
        [BsonId]
        [Browsable(false)]
        public int _id { get; set; }
        public int MaxClipSizeKB { get; set; } = 2000;
        public int HistoryDepthNumber { get; set; } = 10000;
        public int HistoryDepthDays { get; set; } = 0;
        public int MaxCellsToCaptureFormattedText { get; set; }
        public int MaxCellsToCaptureImage { get; set; }
        public int WindowPositionX { get; set; }
        public int WindowPositionY { get; set; }
        public int dataGridViewWidth { get; set; }
        public int MainWindowSizeWidth { get; set; }
        public int MainWindowSizeHeight { get; set; }
        public string DatabaseSize { get; set; }
        public string NumberOfClips { get; set; }
        public string GlobalHotkeyOpenCurrent { get; set; } = "Control + Oemtilde";
        public string GlobalHotkeyOpenLast { get; set; } = "No";
        public string GlobalHotkeyOpenFavorites { get; set; } = "Alt + B";
        public string GlobalHotkeyIncrementalPaste { get; set; } = "No";
        public string GlobalHotkeyDecrementalPaste { get; set; } = "No";
        public string GlobalHotkeyCompareLastClips { get; set; } = "No";
        public string GlobalHotkeyPasteText { get; set; } = "No";
        public string GlobalHotkeySimulateInput { get; set; } = "No";
        public string GlobalHotkeySwitchMonitoring { get; set; } = "No";
        public string GlobalHotkeyForcedCapture { get; set; } = "No";
        public string DefaultFont { get; set; } = "Segoe UI";
        public string Language { get; set; } = "Default";
        public string ClipTempFileFolder { get; set; }
        public string TextCompareApplication { get; set; }
        public string TextEditor { get; set; }
        public string HtmlEditor { get; set; }
        public string RtfEditor { get; set; }
        public string ImageEditor { get; set; }
        public string DatabasePath { get; set; } = "clippy.litedb";
        public string DelimiterForTextJoin { get; set; } = "\n";
        public List<string> IgnoreApplicationsClipCapture { get; set; } = new List<string>();
        public List<string> LastFilterValues { get; set; } = new List<string>();
        public bool IgnoreExclusiveFormatClipCapture { get; set; } = false;
        public bool ShowApplicationIconColumn { get; set; } = true;
        public bool ShowSecondaryColumns { get; set; } = false;
        public bool FastWindowOpen { get; set; } = true;
        public bool MoveCopiedClipToTop { get; set; } = false;
        public bool WindowAutoPosition { get; set; } = true;
        public bool AutoSelectMatch { get; set; } = false;
        public bool SearchWildcards { get; set; } = false;
        public bool SearchWordsIndependently { get; set; } = false;
        public bool SearchCaseSensitive { get; set; } = false;
        public bool SearchIgnoreBigTexts { get; set; } = false;
        public bool FilterListBySearchString { get; set; } = false;
        public bool AutoSelectMatchedClip { get; set; } = false;
        public bool SearchAllFields { get; set; } = true;
        public bool AllowDownloadThumbnail { get; set; } = false;
        public bool ConfirmationBeforeDelete { get; set; } = false;
        public bool MonitoringClipboard { get; set; } = true;
        public bool RestoreCaretPositionOnFocusReturn { get; set; } = false;
        public bool ShowInTaskBar { get; set; } = false;
        public bool Autostart { get; set; } = false;
        public bool UseFormattingInDuplicateDetection { get; set; } = false;
        public bool ReplaceDuplicates { get; set; } = true;
        public bool DeleteNonFavoriteClipsOnExit { get; set; } = false;
        public bool CaptureImages { get; set; } = true;
        public bool DontSendPaste { get; set; } = false;
        public bool ShowNativeTextFormatting { get; set; } = false;
        public bool WordWrap { get; set; } = false;
        public bool MonospacedFont { get; set; } = false;
        //public bool ReadWindowTitles { get; set; } = true;
        public string richTextBoxFontFamily { get; set; } = "Segoe UI";
        public int richTextBoxFontSize { get; set; } = 9;
        public string dataGridViewFontFamily { get; set; } = "Segoe UI";
        public int dataGridViewFontSize { get; set; } = 9;



    }
}




