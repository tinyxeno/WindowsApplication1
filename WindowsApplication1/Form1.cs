using System;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;

namespace WindowsApplication1
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// �t�@�C���t�B���^
        /// </summary>
        private string m_csFilter =
            "Excel 97-2003 �u�b�N (*.xls)|*.xls|" +
            Properties.Resources.AFX_IDS_ALLFILTER +
            "|*.*";

        /// <summary>
        /// �ҏW���t���O
        /// </summary>
        private bool    m_bDirty;

        /// <summary>
        /// �f�[�^�x�[�X�̃e�[�u��
        /// </summary>
        private DataTable m_table;

        /// <summary>
        /// �f�[�^�x�[�X�̍s
        /// </summary>
        private DataRow m_row;

        /// <summary>
        /// �}�E�X�{�^���̉����ʒu
        /// </summary>
        Point m_posPosition;

        /// <summary>
        /// �����_�C�A���O
        /// </summary>
        private Find1 m_find1;

        /// <summary>
        /// �u���_�C�A���O
        /// </summary>
        private Replace1 m_replace1;

        /// <summary>
        /// �R���X�g���N�^
        /// </summary>
        public Form1()
        {
            m_bDirty = false;
            m_table = null;
            m_row = null;
            m_find1 = null;
            m_replace1 = null;
            InitializeComponent();
        }

        /// <summary>
        /// �^�C�g���擾
        /// </summary>
        private string AssemblyTitle
        {
            get
            {
                // ���̃A�Z���u����̃^�C�g�����������ׂĎ擾���܂�
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                // ���Ȃ��Ƃ� 1 �̃^�C�g������������ꍇ
                if (0 < attributes.Length)
                {
                    // �ŏ��̍��ڂ�I�����܂�
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    // ��̕�����̏ꍇ�A���̍��ڂ�Ԃ��܂�
                    if (titleAttribute.Title != string.Empty)
                    {
                        return titleAttribute.Title;
                    }
                }
                // �^�C�g���������Ȃ����A�܂��̓^�C�g����������̕�����̏ꍇ�A.exe ����Ԃ��܂�
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        /// <summary>
        /// �s�ʒu
        /// </summary>
        private int Row
        {
            get
            {
                return listView1.Row;
            }
            set
            {
                if (2 < statusStrip1.Items.Count)
                {
                    bool result = 0 < listView1.Items.Count;
                    // �s�ʒu�̕\��
                    ToolStripStatusLabel label =
                        (ToolStripStatusLabel)statusStrip1.Items[1];
                    if (null != label)
                    {
                        int index = 0;
                        if (result)
                        {
                            index = 1 + value;
                        }
                        label.Text = string.Format(Properties.Resources.ID_INDICATOR_ROW, index);
                        label.Enabled = result;
                    }
                }
                listView1.Row = value;
            }
        }

        /// <summary>
        /// ��ʒu
        /// </summary>
        private int Col
        {
            get
            {
                return listView1.Col;
            }
            set
            {
                if (3 < statusStrip1.Items.Count)
                {
                    bool result = 0 < listView1.Items.Count;
                    // ��ʒu�̕\��
                    ToolStripStatusLabel label =
                        (ToolStripStatusLabel)statusStrip1.Items[2];
                    if (null != label)
                    {
                        int index = 0;
                        if (result)
                        {
                            index = 1 + value;
                        }
                        label.Text = string.Format(Properties.Resources.ID_INDICATOR_COL, index);
                        label.Enabled = result;
                    }
                }
                listView1.Col = value;
            }
        }

        /// <summary>
        /// �t�H�[�����������[�֓W�J���悤�Ƃ����ꍇ�̃C�x���g
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            // ���C���[�W���X�g
            System.Drawing.Icon[] icons = 
                {
                    System.Drawing.SystemIcons.Shield,
                    System.Drawing.SystemIcons.Error,
                    System.Drawing.SystemIcons.Question,
                    System.Drawing.SystemIcons.Warning,
                    System.Drawing.SystemIcons.Information,
                };
            imageList1.ImageSize = new Size(16, 15);
            imageList2.ImageSize = new Size(32, 32);
            foreach (System.Drawing.Icon icon in icons)
            {
                imageList1.Images.Add(icon);
                imageList2.Images.Add(icon);
            };

            // �����j���[�o�[
            newToolStripMenuItem        .ShortcutKeys = Keys.Control | Keys.N;
            openToolStripMenuItem       .ShortcutKeys = Keys.Control | Keys.O;
            saveToolStripMenuItem       .ShortcutKeys = Keys.Control | Keys.S;
            printToolStripMenuItem      .ShortcutKeys = Keys.Control | Keys.P;
            undoToolStripMenuItem       .ShortcutKeys = Keys.Control | Keys.Z;
            cutToolStripMenuItem        .ShortcutKeys = Keys.Control | Keys.X;
            copyToolStripMenuItem       .ShortcutKeys = Keys.Control | Keys.C;
            pasteToolStripMenuItem      .ShortcutKeys = Keys.Control | Keys.V;
            selectallToolStripMenuItem  .ShortcutKeys = Keys.Control | Keys.A;
            findToolStripMenuItem       .ShortcutKeys = Keys.Control | Keys.F;
            repeatToolStripMenuItem     .ShortcutKeys = Keys.F3;
            replaceToolStripMenuItem    .ShortcutKeys = Keys.Control | Keys.H;
            insertToolStripMenuItem     .ShortcutKeys = Keys.Insert;
            propertiesToolStripMenuItem .ShortcutKeys = Keys.F2;
            clearToolStripMenuItem      .ShortcutKeys = Keys.Delete;
            aboutToolStripMenuItem      .ShortcutKeys = Keys.Alt | Keys.OemQuestion;
            aboutToolStripMenuItem      .ShortcutKeyDisplayString = "Alt+?";

            fileToolStripMenuItem       .Text = Properties.Resources.IDM_FILE;
            newToolStripMenuItem        .Text = Properties.Resources.IDM_FILE_NEW;
            openToolStripMenuItem       .Text = Properties.Resources.IDM_FILE_OPEN;
            saveToolStripMenuItem       .Text = Properties.Resources.IDM_FILE_SAVE;
            saveasToolStripMenuItem     .Text = Properties.Resources.IDM_FILE_SAVE_AS;
            setupToolStripMenuItem      .Text = Properties.Resources.IDM_FILE_PRINT_SETUP;
            pageToolStripMenuItem       .Text = Properties.Resources.IDM_FILE_PAGE_SETUP;
            previewToolStripMenuItem    .Text = Properties.Resources.IDM_FILE_PRINT_PREVIEW;
            printToolStripMenuItem      .Text = Properties.Resources.IDM_FILE_PRINT;
            sendtoToolStripMenuItem     .Text = Properties.Resources.IDM_FILE_SEND_MAIL;
            exitToolStripMenuItem       .Text = Properties.Resources.IDM_APP_EXIT;
            editToolStripMenuItem       .Text = Properties.Resources.IDM_EDIT;
            undoToolStripMenuItem       .Text = Properties.Resources.IDM_EDIT_UNDO;
            cutToolStripMenuItem        .Text = Properties.Resources.IDM_EDIT_CUT;
            copyToolStripMenuItem       .Text = Properties.Resources.IDM_EDIT_COPY;
            pasteToolStripMenuItem      .Text = Properties.Resources.IDM_EDIT_PASTE;
            selectallToolStripMenuItem  .Text = Properties.Resources.IDM_EDIT_SELECT_ALL;
            findToolStripMenuItem       .Text = Properties.Resources.IDM_EDIT_FIND;
            repeatToolStripMenuItem     .Text = Properties.Resources.IDM_EDIT_REPEAT;
            replaceToolStripMenuItem    .Text = Properties.Resources.IDM_EDIT_REPLACE;
            insertToolStripMenuItem     .Text = Properties.Resources.IDM_OLE_INSERT_NEW;
            propertiesToolStripMenuItem .Text = Properties.Resources.IDM_OLE_EDIT_PROPERTIES;
            clearToolStripMenuItem      .Text = Properties.Resources.IDM_EDIT_CLEAR;
            viewToolStripMenuItem       .Text = Properties.Resources.IDM_VIEW;
            toolbarToolStripMenuItem    .Text = Properties.Resources.IDM_VIEW_TOOLBAR;
            statusbarToolStripMenuItem  .Text = Properties.Resources.IDM_VIEW_STATUS_BAR;
            largeiconToolStripMenuItem  .Text = Properties.Resources.IDM_VIEW_LARGEICON;
            smalliconToolStripMenuItem  .Text = Properties.Resources.IDM_VIEW_SMALLICON;
            listToolStripMenuItem       .Text = Properties.Resources.IDM_VIEW_LIST;
            detailsToolStripMenuItem    .Text = Properties.Resources.IDM_VIEW_DETAILS;
            helpToolStripMenuItem       .Text = Properties.Resources.IDM_HELP;
            contentsToolStripMenuItem   .Text = Properties.Resources.IDM_HELP_CONTENTS;
            finderToolStripMenuItem     .Text = Properties.Resources.IDM_HELP_FINDER;
            indexToolStripMenuItem      .Text = Properties.Resources.IDM_HELP_INDEX;
            aboutToolStripMenuItem      .Text = Properties.Resources.IDM_APP_ABOUT;

            // ���v���r���[�@�c�[���o�[
            toolStrip1.Visible = false;
            // �u����v���r���[�v�c�[���o�[�́u����v�ݒ�
            SetToolBar(toolStrip1,
                Properties.Resources.ID_PREVIEW_PRINT,
                Properties.Resources.IDB_PREVIEW_PRINT,
                toolStripButton1_Click);
            // �u����v���r���[�v�c�[���o�[�́u����ݒ�v�ݒ�
            SetToolBar(toolStrip1,
                Properties.Resources.ID_PREVIEW_SETUP,
                Properties.Resources.IDB_PREVIEW_SETUP,
                toolStripButton2_Click);
            // �u����v���r���[�v�c�[���o�[�́u�O�ցv�ݒ�
            SetToolBar(toolStrip1,
                Properties.Resources.ID_PREVIEW_PREV,
                Properties.Resources.IDB_PREVIEW_PREV,
                toolStripButton3_Click);
            // �u����v���r���[�v�c�[���o�[�́u���ցv�ݒ�
            SetToolBar(toolStrip1,
                Properties.Resources.ID_PREVIEW_NEXT,
                Properties.Resources.IDB_PREVIEW_NEXT,
                toolStripButton4_Click);
            // �u����v���r���[�v�c�[���o�[�́u�y�[�W�؂�ւ��v�ݒ�
            SetToolBar(toolStrip1,
                Properties.Resources.ID_PREVIEW_ONEPAGE,
                Properties.Resources.IDB_PREVIEW_ONEPAGE,
                toolStripButton5_Click);
            // �u����v���r���[�v�c�[���o�[�́u�g��v�ݒ�
            SetToolBar(toolStrip1,
                Properties.Resources.ID_PREVIEW_ZOOMIN,
                Properties.Resources.IDB_PREVIEW_ZOOMIN,
                toolStripButton6_Click);
            // �u����v���r���[�v�c�[���o�[�́u�k���v�ݒ�
            SetToolBar(toolStrip1,
                Properties.Resources.ID_PREVIEW_ZOOMOUT,
                Properties.Resources.IDB_PREVIEW_ZOOMOUT,
                toolStripButton7_Click);
            // �u����v���r���[�v�c�[���o�[�́u����v�ݒ�
            SetToolBar(toolStrip1,
                Properties.Resources.ID_PREVIEW_CLOSE,
                Properties.Resources.IDB_PREVIEW_CLOSE,
                toolStripButton8_Click);

            // ���W���c�[���o�[
            toolbarToolStripMenuItem_Click(sender, null);

            // �u�W���v�c�[���o�[�́u�V�K�쐬�v�ݒ�
            SetToolBar(toolStrip2,
                Properties.Resources.ID_FILE_NEW,
                Properties.Resources.IDB_FILE_NEW,
                newToolStripMenuItem_Click,
                newToolStripMenuItem);
            // �u�W���v�c�[���o�[�́u�J���v�ݒ�
            SetToolBar(toolStrip2,
                Properties.Resources.ID_FILE_OPEN,
                Properties.Resources.IDB_FILE_OPEN,
                openToolStripMenuItem_Click,
                openToolStripMenuItem);
            // �u�W���v�c�[���o�[�́u�ۑ��v�ݒ�
            SetToolBar(toolStrip2,
                Properties.Resources.ID_FILE_SAVE,
                Properties.Resources.IDB_FILE_SAVE,
                saveToolStripMenuItem_Click,
                saveToolStripMenuItem);
            toolStrip2.Items.Add(new ToolStripSeparator());
            // �u�W���v�c�[���o�[�́u�؂���v�ݒ�
            SetToolBar(toolStrip2,
                Properties.Resources.ID_EDIT_CUT,
                Properties.Resources.IDB_EDIT_CUT,
                cutToolStripMenuItem_Click,
                cutToolStripMenuItem);
            // �u�W���v�c�[���o�[�́u�R�s�[�v�ݒ�
            SetToolBar(toolStrip2,
                Properties.Resources.ID_EDIT_COPY,
                Properties.Resources.IDB_EDIT_COPY,
                copyToolStripMenuItem_Click,
                copyToolStripMenuItem);
            // �u�W���v�c�[���o�[�́u�\��t���v�ݒ�
            SetToolBar(toolStrip2,
                Properties.Resources.ID_EDIT_PASTE,
                Properties.Resources.IDB_EDIT_PASTE,
                pasteToolStripMenuItem_Click,
                pasteToolStripMenuItem);
            toolStrip2.Items.Add(new ToolStripSeparator());
            // �u�W���v�c�[���o�[�́u����v�ݒ�
            SetToolBar(toolStrip2,
                Properties.Resources.ID_FILE_PRINT,
                Properties.Resources.IDB_FILE_PRINT,
                printToolStripMenuItem_Click,
                printToolStripMenuItem);
            toolStrip2.Items.Add(new ToolStripSeparator());
            // �u�W���v�c�[���o�[�́u�傫���A�C�R���v�ݒ�
            SetToolBar(toolStrip2,
                Properties.Resources.ID_VIEW_LARGEICON,
                Properties.Resources.IDB_VIEW_LARGEICON,
                largeiconToolStripMenuItem_Click,
                largeiconToolStripMenuItem);
            // �u�W���v�c�[���o�[�́u�������A�C�R���v�ݒ�
            SetToolBar(toolStrip2,
                Properties.Resources.ID_VIEW_SMALLICON,
                Properties.Resources.IDB_VIEW_SMALLICON,
                smalliconToolStripMenuItem_Click,
                smalliconToolStripMenuItem);
            // �u�W���v�c�[���o�[�́u�ꗗ�v�ݒ�
            SetToolBar(toolStrip2,
                Properties.Resources.ID_VIEW_LIST,
                Properties.Resources.IDB_VIEW_LIST,
                listToolStripMenuItem_Click,
                listToolStripMenuItem);
            // �u�W���v�c�[���o�[�́u�ڍׁv�ݒ�
            SetToolBar(toolStrip2,
                Properties.Resources.ID_VIEW_DETAILS,
                Properties.Resources.IDB_VIEW_DETAILS,
                detailsToolStripMenuItem_Click,
                detailsToolStripMenuItem);
            toolStrip2.Items.Add(new ToolStripSeparator());
            // �u�W���v�c�[���o�[�́u�o�[�W�������v�ݒ�
            SetToolBar(toolStrip2,
                Properties.Resources.ID_APP_ABOUT,
                Properties.Resources.IDB_APP_ABOUT,
                aboutToolStripMenuItem_Click,
                aboutToolStripMenuItem);

            // ���X�e�[�^�X�o�[
            statusbarToolStripMenuItem_Click(sender, null);
            ToolStripStatusLabel item =
                (ToolStripStatusLabel)statusStrip1.Items.Add(string.Empty);
            if (null != item)
            {
                item.Spring = true;
                item.TextAlign = ContentAlignment.MiddleLeft;
                item.ImageAlign = ContentAlignment.MiddleLeft;
                item.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
                item.TextImageRelation = TextImageRelation.ImageBeforeText;
            }
            statusStrip1.Items.Add(Properties.Resources.ID_INDICATOR_ROW);
            statusStrip1.Items.Add(Properties.Resources.ID_INDICATOR_COL);
            statusStrip1.Items.Add(Properties.Resources.ID_INDICATOR_CAPS);
            statusStrip1.Items.Add(Properties.Resources.ID_INDICATOR_NUM);
            statusStrip1.Items.Add(Properties.Resources.ID_INDICATOR_SCRL);
            statusStrip1.Items.Add(Properties.Resources.ID_INDICATOR_MODIFY);
            statusStrip1.Items.Add(Properties.Resources.ID_INDICATOR_DATE);
            statusStrip1.Items.Add(Properties.Resources.ID_INDICATOR_TIME);

            // ���v�����^�h�L�������g
            printDialog1.Document = printDocument1;
            pageSetupDialog1.Document = printDocument1;
            printPreviewControl1.Document = printDocument1;

            // �����X�g�r���[
            listView1.AutoArrange           = true;
            listView1.AllowColumnReorder    = true;
            listView1.BackgroundImageTiled  = true;
            listView1.FullRowSelect         = true;
            listView1.HeaderStyle           = ColumnHeaderStyle.Clickable;
            listView1.HideSelection         = false;
            listView1.HotTracking           = false;
            listView1.HoverSelection        = false;
            listView1.LabelEdit             = true;
            listView1.LargeImageList        = imageList2;
            listView1.MultiSelect           = false;
            listView1.SmallImageList        = imageList1;
            listView1.SubItemImages         = true;
            listView1.ListViewItemSorter    = new ListViewColumnSorter();
            // �s��ʒu�̏�����
            Row = Col = 0;
            // ���X�g�r���[�A�C�e���I�u�W�F�N�g�̕\���ݒ�B�ڍו\���ݒ�
            detailsToolStripMenuItem_Click(sender, e);

            // ���h�L�������g�̏�����
            newToolStripMenuItem_Click(sender, e);

            // ���^�C�}�[�֘A
            timer1.Interval = 125;
            timer1.Enabled = true;
            timer2.Interval = 500;
            timer2.Enabled = false;
        }

        /// <summary>
        /// �t�H�[������悤�Ƃ��Ă���ꍇ�̃C�x���g
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            if (menuStrip1.Visible)
            {
                bool result = false;
                if (m_bDirty)
                {
                    result = SaveModified();
                    if (result)
                    {
                        m_bDirty = false;
                    }
                }
                else
                {
                    switch (AfxMessageBox(Properties.Resources.AFX_IDP_ASK_TO_EXIT,
                        (int)MessageBoxButtons.YesNo +
                        (int)MessageBoxIcon.Question +
                        (int)MessageBoxDefaultButton.Button2))
                    {
                        case DialogResult.Yes:
                            result = true;
                            break;
                    }
                }
                if (result)
                {
                    // �^�C�}�[��~
                    timer1.Enabled = false;
                    timer2.Enabled = false;

                    // �����_�C�A���O���J���Ă���΁A����
                    if (null != m_find1)
                    {
                        m_find1.Close();
                        m_find1 = null;
                    }
                    // �u���_�C�A���O���J���Ă���΁A����
                    if (null != m_replace1)
                    {
                        m_replace1.Close();
                        m_replace1 = null;
                    }
                }
                e.Cancel = !result;
            }
            else
            {
                menuStrip1.Visible = true;
                toolStrip1.Visible = false;
                toolStrip2.Visible = true;
                statusStrip1.Visible = true;
                listView1.Visible = true;
                Form1_Resize(sender, null);
            }
        }

        /// <summary>
        /// �t�H�[����傫����ύX����ꍇ�̏���
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Resize(object sender, EventArgs e)
        {
            Rectangle rectangle = ClientRectangle;
            if (statusStrip1.Visible)
            {
                rectangle.Height -= statusStrip1.Height;
            }
            if (menuStrip1.Visible)
            {
                rectangle.Y += menuStrip1.Height;
                rectangle.Height -= menuStrip1.Height;
                if (toolStrip2.Visible)
                {
                    rectangle.Y += toolStrip2.Height;
                    rectangle.Height -= toolStrip2.Height;
                }
                listView1.Location = rectangle.Location;
                listView1.Size = rectangle.Size;
            }
            else
            {
                if (toolStrip1.Visible)
                {
                    rectangle.Y += toolStrip1.Height;
                    rectangle.Height -= toolStrip1.Height;
                }
                printPreviewControl1.Location = rectangle.Location;
                printPreviewControl1.Size = rectangle.Size;
            }
        }

        /// <summary>
        /// �C���^�[�o���^�C�}�[�̏���
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            // ���X�g�r���[�̃t���[�e�B���O�E�J�[�\���̏�Ԃŋ���/�֎~����
            // �u���ɖ߂��v
            undoToolStripMenuItem.Enabled = listView1.Undo;
            // �u�؂���v
            toolStrip2.Items[4]         .Enabled =
            cutToolStripMenuItem        .Enabled = /*listView1.Cut;*/
            // �u�R�s�[�v
            toolStrip2.Items[5]         .Enabled =
            copyToolStripMenuItem       .Enabled = listView1.Copy;
            // �u�\��t���v
            toolStrip2.Items[6]         .Enabled =
            pasteToolStripMenuItem      .Enabled = listView1.Paste;
            // �u���ׂđI���v
            selectallToolStripMenuItem  .Enabled = listView1.SelectAll;

            // �����\���ƃ|�[�����O�Ɉˑ�����\��
            DateTime datetime = DateTime.Now;
            bool result = !string.IsNullOrEmpty(openFileDialog1.FileName);
            // �uCAPS�v
            statusStrip1.Items[3].Enabled = Control.IsKeyLocked(Keys.Capital);
            // �uNUM�v
            statusStrip1.Items[4].Enabled = Control.IsKeyLocked(Keys.NumLock);
            // �uSCRL�v
            statusStrip1.Items[5].Enabled = Control.IsKeyLocked(Keys.Scroll);
            // �u�ҏW�ς݁v
            statusStrip1.Items[6].Enabled = m_bDirty;
            // �u���t�v
            statusStrip1.Items[7].Text = datetime.ToString(Properties.Resources.ID_INDICATOR_DATE);
            // �u�����v
            string[] source = Properties.Resources.ID_INDICATOR_TIME.Replace("\\n", "\n").Split('\n');
            string value = source[0];
            if (datetime.Millisecond < 500 && 1 < source.Length)
            {
                value = source[1];
            }
            statusStrip1.Items[8].Text = datetime.ToString(value);
        }

        /// <summary>
        /// �f�B���C�E�����V���b�g�E�^�C�}�[�̏���
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer2_Tick(object sender, EventArgs e)
        {
            timer2.Enabled = false;
            propertiesToolStripMenuItem_Click(sender, e);
        }

        /// <summary>
        /// �u�V�K�쐬�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (SaveModified())
            {
                m_bDirty = false;
                openFileDialog1.FileName = string.Empty;
                OnInitialUpdate();
                AfxMessageBox(Properties.Resources.AFX_IDS_IDLEMESSAGE,
                    (int)MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// �h�L�������g���u�J���v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = m_csFilter;
            switch (openFileDialog1.ShowDialog(this))
            {
                case DialogResult.OK:
                    m_bDirty = false;
                    OnInitialUpdate();
                    AfxMessageBox(Properties.Resources.AFX_IDS_IDLEMESSAGE,
                        (int)MessageBoxIcon.Information);
                    break;
                default:
                    AfxMessageBox(Properties.Resources.IDS_AFXBARRES_CANCEL,
                        (int)MessageBoxIcon.Information);
                    break;
            }
        }

        /// <summary>
        /// �h�L�������g���u�ۑ�����v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(openFileDialog1.FileName))
            {
                // TODO: �t�@�C���̕ۑ�
                m_bDirty= false;
                AfxMessageBox(Properties.Resources.AFX_IDS_IDLEMESSAGE,
                    (int)MessageBoxIcon.Information);
            }
            else
            {
                // �t�@�C�������w�肳��Ă��Ȃ��̂ŁA�u���O�����ĕۑ��v���N��
                saveasToolStripMenuItem_Click(sender, e);
            }
        }

        /// <summary>
        /// �h�L�������g���u���O�����ĕۑ�����v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = m_csFilter;
            saveFileDialog1.FileName = Properties.Resources.AFX_IDS_UNTITLED;
            if (!string.IsNullOrEmpty(openFileDialog1.FileName))
            {
                saveFileDialog1.FileName = openFileDialog1.FileName;
            }
            switch (saveFileDialog1.ShowDialog(this))
            {
                case DialogResult.OK:
                    // �t�@�C������ێ�����
                    openFileDialog1.FileName = saveFileDialog1.FileName;
                    // �t�@�C�����̃^�C�g�����t�H�[���̃^�C�g���֔��f����
                    SetTitle();
                    // �u�ۑ��v���N��
                    saveToolStripMenuItem_Click(sender, e);
                    break;
                default:
                    AfxMessageBox(Properties.Resources.IDS_AFXBARRES_CANCEL,
                        (int)MessageBoxIcon.Information);
                    break;
            }
        }

        /// <summary>
        /// �u����ݒ�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void setupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.ShowPrintSetupDlg();
        }

        /// <summary>
        /// �u�y�[�W�ݒ�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void pageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            switch (pageSetupDialog1.ShowDialog(this))
            {
                case DialogResult.OK:
                    break;
                default:
                    AfxMessageBox(Properties.Resources.IDS_AFXBARRES_CANCEL,
                        (int)MessageBoxIcon.Information);
                    break;
            }
        }

        /// <summary>
        /// �u����v���r���[�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void previewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // �����_�C�A���O���J���Ă���΁A����
            if (null != m_find1)
            {
                m_find1.Close();
                m_find1 = null;
            }
            // �u���_�C�A���O���J���Ă���΁A����
            if (null != m_replace1)
            {
                m_replace1.Close();
                m_replace1 = null;
            }
            menuStrip1.Visible = false;
            toolStrip1.Visible = true;
            toolStrip2.Visible = false ;
            statusStrip1.Visible = true;
            listView1.Visible = false;
            Form1_Resize(sender, e);
        }

        /// <summary>
        /// �u����v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            switch (printDialog1.ShowDialog(this))
            {
                case DialogResult.OK:
                    break;
                default:
                    AfxMessageBox(Properties.Resources.IDS_AFXBARRES_CANCEL,
                        (int)MessageBoxIcon.Information);
                    break;
            }
        }

        /// <summary>
        /// �u���M�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void sendtoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // ���ʉ������Ȃ�
        }

        /// <summary>
        /// �u�I���v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// �u���ɖ߂��v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.Undo = true;
        }

        /// <summary>
        /// �u�؂���v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.Cut = true;
        }

        /// <summary>
        /// �u�R�s�[�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.Copy = true;
        }

        /// <summary>
        /// �u�\��t���v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.Paste = true;
        }

        /// <summary>
        /// �u���ׂđI���v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void selectallToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.SelectAll = true;
        }

        /// <summary>
        /// �u�����v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (null == m_find1 && null == m_replace1)
            {
                m_find1 = new Find1();
                if (null != m_find1)
                {
                    m_find1.m_eventHandler += findReplaceHandler;
                    m_find1.Show(this);
                }
            }
        }

        /// <summary>
        /// �u���ցv����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void repeatToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // TODO:
        }

        /// <summary>
        /// �u�u���v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void replaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (null == m_find1 && null == m_replace1)
            {
                m_replace1 = new Replace1();
                if (null != m_replace1)
                {
                    m_replace1.m_eventHandler += findReplaceHandler;
                    m_replace1.Show(this);
                }
            }
        }

        /// <summary>
        /// �u�}���v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void insertToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // TODO:
        }

        /// <summary>
        /// �u�ҏW�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void propertiesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.Items[Row].BeginEdit();
        }

        /// <summary>
        /// �u�폜�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // TODO:
        }

        /// <summary>
        /// �u�c�[�� �o�[�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolbarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool result = toolStrip2.Visible;
            if (null != e)
            {
                result = !result;
            }
            toolStrip2.Visible = result;
            toolbarToolStripMenuItem.Checked = result;
            Form1_Resize(sender, e);
        }

        /// <summary>
        /// �u�X�e�[�^�X �o�[�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void statusbarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool result = statusStrip1.Visible;
            if (null != e)
            {
                result = !result;
            }
            statusStrip1.Visible = result;
            statusbarToolStripMenuItem.Checked = result;
            Form1_Resize(sender, e);
        }

        /// <summary>
        /// �u�傫���A�C�R���v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void largeiconToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OnChangeView(View.LargeIcon);
        }

        /// <summary>
        /// �u�������A�C�R���v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void smalliconToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OnChangeView(View.SmallIcon);
        }

        /// <summary>
        /// �u�ꗗ�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OnChangeView(View.List);
        }

        /// <summary>
        /// �u�ڍׁv����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void detailsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OnChangeView(View.Details);
        }

        /// <summary>
        /// �u�ڎ��v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void contentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Help.ShowHelp(this, Properties.Resources.IDS_HELP_FILENAME);
        }

        /// <summary>
        /// �u�w���v�����v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void finderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Help.ShowHelp(this, Properties.Resources.IDS_HELP_FILENAME, "");
        }

        /// <summary>
        /// �u�����v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void indexToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Help.ShowHelpIndex(this, Properties.Resources.IDS_HELP_FILENAME);
        }

        /// <summary>
        /// �u�o�[�W�������v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 aboutBox1 = new AboutBox1();
            if (null != aboutBox1)
            {
                aboutBox1.Show(this);
            }
        }

        /// <summary>
        /// ����v���r���[�́u����v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Close();
            printToolStripMenuItem_Click(sender, e);
        }

        /// <summary>
        /// ����v���r���[�́u�ݒ�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Close();
            setupToolStripMenuItem_Click(sender, e);
        }

        /// <summary>
        /// ����v���r���[�́u�O�ցv����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            // TODO:
        }

        /// <summary>
        /// ����v���r���[�́u���ցv����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            // TODO:
        }

        /// <summary>
        /// ����v���r���[�́u1�y�[�W/2�y�[�W�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            switch (printPreviewControl1.Columns)
            {
                case 1:
                    printPreviewControl1.Columns = 2;
                    toolStrip1.Items[4].Text = Properties.Resources.AFX_IDS_ONEPAGE;
                    break;
                default:
                    printPreviewControl1.Columns = 1;
                    toolStrip1.Items[4].Text = Properties.Resources.AFX_IDS_TWOPAGE;
                    break;
            }
        }

        /// <summary>
        /// ����v���r���[�́u�g��v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            bool result = true;
            double value = printPreviewControl1.Zoom;
            if (value < 0.2)
            {
                value = 0.2;
            }
            else if (value < 0.4)
            {
                value = 0.4;
            }
            else if (value < 0.6)
            {
                value = 0.6;
            }
            else if (value < 0.8)
            {
                value = 0.8;
            }
            else if (value < 1.0)
            {
                value = 1.0;
            }
            else if (value < 1.2)
            {
                value = 1.2;
            }
            else if (value < 1.4)
            {
                value = 1.4;
            }
            else if (value < 1.6)
            {
                value = 1.6;
            }
            else if (value < 1.8)
            {
                value = 1.8;
                result = false;
            }
            printPreviewControl1.Zoom = value;
            toolStrip1.Items[5].Enabled = result;
            toolStrip1.Items[6].Enabled = true;
        }

        /// <summary>
        /// ����v���r���[�́u�k���v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            bool result = true;
            double value = printPreviewControl1.Zoom;
            if (1.6 < value)
            {
                value = 1.6;
            }
            else if (1.4 < value)
            {
                value = 1.4;
            }
            else if (1.2 < value)
            {
                value = 1.2;
            }
            else if (1.0 < value)
            {
                value = 1.0;
            }
            else if (0.8 < value)
            {
                value = 0.8;
            }
            else if (0.6 < value)
            {
                value = 0.6;
            }
            else if (0.4 < value)
            {
                value = 0.4;
            }
            else if (0.2 < value)
            {
                value = 0.2;
                result = false;
            }
            printPreviewControl1.Zoom = value;
            toolStrip1.Items[5].Enabled = true;
            toolStrip2.Items[6].Enabled = result;
        }

        /// <summary>
        /// ����v���r���[�́u�I���v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// �������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            // TODO:
        }

        /// <summary>
        /// ���X�g�r���[�́u�}�E�X�{�^�������v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_MouseDown(object sender, MouseEventArgs e)
        {
            m_posPosition = e.Location;
        }

        /// <summary>
        /// ���X�g�r���[�́u�N���b�N�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_Click(object sender, EventArgs e)
        {
            ListViewHitTestInfo info = listView1.HitTest(m_posPosition);
            if (null != info)
            {
                Row = info.Item.Index;
                Col = info.Item.SubItems.IndexOf(info.SubItem);
            }
        }

        /// <summary>
        /// ���X�g�r���[�́u�A�C�e���I��ύX�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            switch (listView1.View)
            {
                case View.Details:
                    break;
                default:
                    Col = 0;
                    break;
            }
            if (Row != e.Item.Index)
            {
                Row = e.Item.Index;
            }
        }

        /// <summary>
        /// ���X�g�r���[�́u�L�[�{�[�h�̃L�[�̉����v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Left:
                case Keys.Right:
                    if (1 < listView1.Items.Count)
                    {
                        int value = listView1.Columns[Col].DisplayIndex;
                        switch (e.KeyCode)
                        {
                            case Keys.Left:
                                if (0 < value)
                                {
                                    value--;
                                }
                                break;
                            case Keys.Right:
                                if (value < (listView1.Columns.Count - 1))
                                {
                                    value++;
                                }
                                break;
                        }
                        if (value != listView1.Columns[Col].DisplayIndex)
                        {
                            bool result = false;
                            for (int index = 0; false == result && index < listView1.Columns.Count; index++)
                            {
                                result = listView1.Columns[index].DisplayIndex == value;
                                if (result)
                                {
                                    Col = index;
                                }
                            };
                        }
                    }
                    break;
            }
        }

        /// <summary>
        /// ���X�g�r���[�́u�J�����w�b�_�̕`��v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            e.DrawDefault = true;
        }

        /// <summary>
        /// ���X�g�r���[�́u�A�C�e���`��v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
            e.DrawDefault = true;
        }

        /// <summary>
        /// ���X�g�r���[�́u�T�u�A�C�e���`��v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            e.DrawDefault = true;
        }

        /// <summary>
        /// ���X�g�r���[�́u�J�����N���b�N�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            ListViewColumnSorter source = (ListViewColumnSorter)listView1.ListViewItemSorter;
            source.SortColumn = e.Column;
            for (int index = 0; index < listView1.Columns.Count; index++)
            {
                SortOrder value = SortOrder.None;
                if (e.Column == index)
                {
                    switch (listView1.GetSortOrder(index))
                    {
                        case SortOrder.Ascending:
                            value = SortOrder.Descending;
                            break;
                        default:
                            value = SortOrder.Ascending;
                            break;
                    }
                    source.Order = value;
                }
                listView1.SetSortOrder(index, value);
            };
            listView1.Sort();
        }

        /// <summary>
        /// ���X�g�r���[�́u�_�u���N���b�N�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            propertiesToolStripMenuItem_Click(sender, e);
        }

        /// <summary>
        /// ���X�g�r���[�́u���x���ҏW�J�n�v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_BeforeLabelEdit(object sender, LabelEditEventArgs e)
        {
            string value = string.Empty;
            if (DDX_ListViewItemText(true, Col, ref value))
            {
                listView1.LabelString = value;
                AfxMessageBox(Properties.Resources.IDP_AFXBARRES_TEXT_IS_REQUIRED,
                    (int)MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// ���X�g�r���[�́u���x���ҏW�I���v����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_AfterLabelEdit(object sender, LabelEditEventArgs e)
        {
            e.CancelEdit = true;
            if (null != e.Label)
            {
                if (!string.IsNullOrEmpty(e.Label))
                {
                    string value = e.Label;
                    if (DDX_ListViewItemText(false, Col, ref value))
                    {
                        m_bDirty = true;
                        AfxMessageBox(Properties.Resources.IDS_EDIT_MENU,
                            (int)MessageBoxIcon.Information);
                    }
                }
                else
                {
                    AfxMessageBox(Properties.Resources.IDP_AFXBARRES_TEXT_IS_REQUIRED);
                    timer2.Enabled = true;
                }
            }
            else
            {
                AfxMessageBox(Properties.Resources.IDS_AFXBARRES_CANCEL,
                    (int)MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="toolStrip">�c�[���o�[</param>
        /// <param name="caption">�ݒ蕶����</param>
        /// <param name="image">�C���[�W</param>
        /// <param name="eventHandler">�C�x���g�n���h��</param>
        /// <param name="menuItem">�y�ȗ��z���j���[�A�C�e��</param>
        private void SetToolBar(ToolStrip toolStrip, string caption, Image image, EventHandler eventHandler) { SetToolBar(toolStrip, caption, image, eventHandler, null); }
        private void SetToolBar(ToolStrip toolStrip, string caption, Image image, EventHandler eventHandler, ToolStripMenuItem menuItem)
        {
            string[] value = caption.Replace("\\n", "\n").Split('\n');
            ToolStripButton source = (ToolStripButton)toolStrip.Items.Add(value[2], image, eventHandler);
            if (null != source)
            {
                source.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
                source.ToolTipText = value[1] + GetKeyString(menuItem) + "\n" + value[0];
                if (null != menuItem)
                {
                    source.TextImageRelation = TextImageRelation.ImageAboveText;
                    menuItem.Image = image;
                }
            }
        }

        /// <summary>
        /// ���b�Z�[�W�{�b�N�X�\������
        /// </summary>
        /// <param name="caption">�\��������</param>
        /// <param name="style">���b�Z�[�W�{�b�N�X�̕\���X�^�C��</param>
        /// <returns>���b�Z�[�W�{�b�N�X�̑��쌋��</returns>
        private DialogResult AfxMessageBox(string caption) { return AfxMessageBox(caption, (int)MessageBoxIcon.Exclamation); }
        private DialogResult AfxMessageBox(string caption, int style)
        {
            int index = 7;
            index &= style;
            MessageBoxButtons button = (MessageBoxButtons)index;
            index = 
                (int)MessageBoxDefaultButton.Button2 |
                (int)MessageBoxDefaultButton.Button3;
            index &= style;
            MessageBoxDefaultButton def_button = (MessageBoxDefaultButton)index;
            index =
                (int)MessageBoxIcon.Exclamation |
                (int)MessageBoxIcon.Information;
            index &= style;
            MessageBoxIcon icon = (MessageBoxIcon)index;

            index >>= 4;
            string value = string.Empty;
            string[] source = Properties.Resources.IDS_MESSAGE_TITLES.Replace("\\n", "\n").Split('\n');
            if (0 <= index && index < source.Length)
            {
                value = source[index];
            }

            if (0 < statusStrip1.Items.Count)
            {
                ToolStripStatusLabel item 
                    = (ToolStripStatusLabel)statusStrip1.Items[0];
                if (null != item)
                {
                    Image image = null;
                    if (0 <= index && index < imageList1.Images.Count)
                    {
                        image = imageList1.Images[index];
                    }
                    if (null != image)
                    {
                        item.Image = image;
                    }
                    item.Text = caption;
                }
            }

            DialogResult result = DialogResult.OK;
            switch (style)
            {
                case (int)MessageBoxIcon.Information:
                    break;
                default:
                    result = MessageBox.Show(this, caption,value, button,icon,def_button);
                    break;
            }
            switch (result)
            {
                case DialogResult.Cancel:
                case DialogResult.No:
                    AfxMessageBox(Properties.Resources.IDS_AFXBARRES_CANCEL, (int)MessageBoxIcon.Information);
                    break;
                case DialogResult.Abort:
                    AfxMessageBox(Properties.Resources.IDS_AFXBARRES_ABORT, (int)MessageBoxIcon.Information);
                    break;
            }
            return result;
        }

        /// <summary>
        /// �h�L�������g�̕ҏW��Ԃ̊m�F����
        /// </summary>
        /// <returns></returns>
        private bool SaveModified()
        {
            bool result = true;
            if (m_bDirty)
            {
                result = false;
                string source = string.Empty;
                source = Properties.Resources.AFX_IDS_UNTITLED;
                if (!string.IsNullOrEmpty(openFileDialog1.FileName))
                {
                    source = Path.GetFileName(openFileDialog1.FileName);
                }
                switch (AfxMessageBox(Properties.Resources.AFX_IDP_ASK_TO_SAVE.Replace("%1", source),
                    (int)MessageBoxButtons.YesNoCancel +
                    (int)MessageBoxIcon.Question +
                    (int)MessageBoxDefaultButton.Button3))
                {
                    case DialogResult.Yes:
                        saveToolStripMenuItem_Click(null, null);
                        result = !m_bDirty;
                        break;
                    case DialogResult.No:
                        result = true;
                        break;
                }
            }
            return result;
        }

        /// <summary>
        /// �t�H�[���փ^�C�g����ݒ肷�鏈��
        /// </summary>
        private void SetTitle()
        {
            string value = Properties.Resources.AFX_IDS_UNTITLED;
            if (!string.IsNullOrEmpty(openFileDialog1.FileName))
            {
                value = Path.GetFileName(openFileDialog1.FileName);
            }
            Text = AssemblyTitle + Properties.Resources.AFX_IDS_OBJ_TITLE_INPLACE + value;
        }

        /// <summary>
        /// ���j���[�A�C�e���ɐݒ肳��Ă���V���[�g�J�b�g���𕶎���ŕԂ�
        /// </summary>
        /// <param name="toolStripMenuItem">���j���[�A�C�e��</param>
        /// <returns>�V���[�g�J�b�g�ɊY�����镶������J�b�R�t���ŕԂ��B
        /// ���j���[�A�C�e���ɃV���[�g�J�b�g���ݒ肳��Ă��Ȃ���΋��Ԃ�</returns>
        private string GetKeyString(ToolStripMenuItem toolStripMenuItem)
        {
            string result = string.Empty;
            if (null != toolStripMenuItem)
            {
                Keys keys = toolStripMenuItem.ShortcutKeys;
                if (Keys.None != keys)
                {
                    if (0 != (keys & Keys.Control))
                    {
                        result += "Ctrl+";
                    }
                    if (0 != (keys & Keys.Shift))
                    {
                        result += "Shift+";
                    }
                    if (0 != (keys & Keys.Alt))
                    {
                        result += "Alt+";
                    }
                    int value = (int)keys;
                    value &= 0xFF;
                    Keys key = (Keys)value;
                    if (Keys.A <= key && key <= Keys.Z)
                    {
                        result += (Char)value;
                    }
                    else if (Keys.F1 <= key && key <= Keys.F24)
                    {
                        result += "F" + (value - (int)Keys.F1 + 1)
                            .ToString();
                    }
                    else
                    {
                        switch (key)
                        {
                            case Keys.OemQuestion:
                                result += "?";
                                break;
                            case Keys.Insert:
                                result += "Ins";
                                break;
                            case Keys.Delete:
                                result += "Del";
                                break;
                        }
                    }
                }
            }
            if (!string.IsNullOrEmpty(result))
            {
                result = " (" + result + ")";
            }
            return result;
        }

        /// <summary>
        /// �r���[�̕\����Ԃ̕ύX����
        /// </summary>
        /// <param name="value"></param>
        private void OnChangeView(View value)
        {
            listView1.View = value;
            ToolStripButton
                button = (ToolStripButton)toolStrip2.Items[10];
                button                  .Checked =
            largeiconToolStripMenuItem  .Checked = View.LargeIcon   == value;
                button = (ToolStripButton)toolStrip2.Items[11];
                button                  .Checked =
            smalliconToolStripMenuItem  .Checked = View.SmallIcon   == value;
                button = (ToolStripButton)toolStrip2.Items[12];
                button                  .Checked =
            listToolStripMenuItem       .Checked = View.List        == value;
                button = (ToolStripButton)toolStrip2.Items[13];
                button                  .Checked =
            detailsToolStripMenuItem    .Checked = View.Details     == value;
        }

        /// <summary>
        /// �r���[�ƃR�}���h�{�^���̍X�V����
        /// </summary>
        private void OnInitialUpdate()
        {
            // �t�H�[���փ^�C�g���̐ݒ肷��
            SetTitle();

            //�r���[�̍X�V����
            OnUpdate();

            Row = Col = 0;
            // �ȉ��A�h�L�������g���J���Ă��邩�ǂ����𔻒�
            bool result = !string.IsNullOrEmpty(openFileDialog1.FileName);

            // �u���M�v�F��ɋ֎~
            sendtoToolStripMenuItem .Enabled = false;

            // �u�J���v�F�h�L�������g�����Ă���ꍇ�ɋ���
            toolStrip2.Items[1]     .Enabled = 
            openToolStripMenuItem   .Enabled = !result;

            // �ȉ��A�h�L�������g���J���Ă���ꍇ�ɋ���
            // �u���X�g�r���[�v
            listView1.BackgroundImage = null;
            if (result)
            {
                listView1.BackgroundImage = Properties.Resources.IDB_VIEW_STRIPED;
            }
            listView1               .GridLines =
            listView1               .Enabled =
            // �u�V�K�쐬�v
            toolStrip2.Items[0]     .Enabled =
            newToolStripMenuItem    .Enabled =
            // �u�ۑ��v
            toolStrip2.Items[2]     .Enabled =
            saveToolStripMenuItem   .Enabled =
            // �u���O�����ĕۑ��v
            saveasToolStripMenuItem .Enabled =
            // �u����v
            toolStrip2.Items[4]     .Enabled =
            printToolStripMenuItem  .Enabled =
            // �u�y�[�W�ݒ�v
            pageToolStripMenuItem   .Enabled =
            // �u����v���r���[�v
            previewToolStripMenuItem.Enabled =
            // �u�}���v
            insertToolStripMenuItem .Enabled = result;

            // �ȉ��A���X�g�r���[�Ƀf�[�^�����邩�ǂ����ŋ��֎~��ݒ�
            SetCursor();
        }

        /// <summary>
        /// �R�}���h�{�^���̋���/�֎~�ƌ��ݒ��ڒ��̍s��ʒu�̕\��
        /// </summary>
        private void SetCursor()
        {
            string value = string.Empty;
            bool result = 0 < listView1.Items.Count;

            // �ȉ��A���X�g�r���[�Ƀf�[�^�����邩�ǂ����ŋ��֎~��ݒ�
            // �u�����v
            findToolStripMenuItem       .Enabled =
            // �u���ցv
            repeatToolStripMenuItem     .Enabled =
            // �u�u���v
            replaceToolStripMenuItem    .Enabled =
            // �u�ҏW�v
            propertiesToolStripMenuItem .Enabled =
            // �u�폜�v
            clearToolStripMenuItem      .Enabled = result;

            // �A�C�e����I�����A�t�H�[�J�X�𓖂Ă�
            if (0 <= Row && Row < listView1.Items.Count)
            {
                ListViewItem item = listView1.Items[Row];
                if (null != item)
                {
                    item.Focused    =
                    item.Selected   = true;
                    item.EnsureVisible();
                }
                listView1.Focus();
            }
        }

        /// <summary>
        /// �r���[�̍X�V����
        /// </summary>
        private void OnUpdate()
        {
            // �\�[�g�����̃N���A
            ListViewColumnSorter source =
                (ListViewColumnSorter)listView1.ListViewItemSorter;
            source.Clear();
            // ���X�g�r���[�A�C�e���̃N���A
            listView1.Items.Clear();
            // ���X�g�r���[�J�����w�b�_�̃N���A
            listView1.Columns.Clear();
            if (!string.IsNullOrEmpty(openFileDialog1.FileName))
            {
                DbProviderFactory factory =
                    DbProviderFactories.GetFactory("System.Data.OleDb");
                using (DbConnection conn = factory.CreateConnection())
                {
                    DbConnectionStringBuilder builder =
                        factory.CreateConnectionStringBuilder();
                    switch (IntPtr.Size)
                    {
                        case 4:
                            builder["Provider"] =
                                Properties.Resources.IDS_CONNECTION_STRING_32;
                            break;
                        case 8:
                            builder["Provider"] =
                                Properties.Resources.IDS_CONNECTION_STRING_64;
                            break;
                    }
                    builder["Data Source"] = openFileDialog1.FileName;
                    builder["Extended Properties"] =
                        Properties.Resources.IDS_CONNECTION_EXTENDED_PROPERTIES;

                    conn.ConnectionString = builder.ToString();
                    conn.Open();
                    using (DbCommand command = conn.CreateCommand())
                    {
                        command.CommandText = Properties.Resources.IDS_SQL_SELECT;
                        m_table = new DataTable();
                        using (DbDataReader reader = command.ExecuteReader())
                        {
                            m_table.Load(reader);
                            // ���X�g�r���[�J�����w�b�_�̍\�z
                            foreach (DataColumn col in m_table.Columns)
                            {
                                listView1.Columns.Add(col.ColumnName, 100, HorizontalAlignment.Left);
                            }
                            // ���X�g�r���[�A�C�e���̍\�z
                            bool result = true;
                            Row = 0;
                            foreach (DataRow row in m_table.Rows)
                            {
                                m_row = row;
                                result = UpdateData(false);
                                if (!result)
                                {
                                    break;
                                }
                                Row++;
                            };
                            m_row = null;
                        }
                        m_table = null;
                    }
                    conn.Close();
                }
            }
        }

        /// <summary>
        /// �r���[�̍X�V/�ǂݏo������
        /// </summary>
        /// <param name="bSave">�y�ȗ��z�r���[�̍X�V����/�ǂݏo���t���O�B�ȗ����A�^�y�ǂݏo�������z</param>
        /// <returns></returns>
        private bool UpdateData() { return UpdateData(true); }
        private bool UpdateData(bool bSave)
        {
            return DoDataExchange(bSave);
        }

        private bool DoDataExchange(bool bSaveAndValidate)
        {
            bool result = true;
            int index = 0;
            foreach (DataColumn col in m_table.Columns)
            {
                string value = m_row[col.ColumnName].ToString();
                result = DDX_ListViewItemText(bSaveAndValidate, index, ref value);
                if (!result)
                {
                    break;
                }
                index++;
            }
            return result;
        }

        /// <summary>
        /// �r���[�̍X�V/�ǂݏo������
        /// </summary>
        /// <param name="bSaveAndValidate">�r���[�̍X�V����/�ǂݏo���t���O�B</param>
        /// <param name="iSubItem">�s�ʒu</param>
        /// <param name="value">�y�Q�Ɓz�ǂݏ������镶����</param>
        /// <returns>�ǂݏ��������̏ꍇ�A�U�ȊO��Ԃ�</returns>
        private bool DDX_ListViewItemText(bool bSaveAndValidate, int iSubItem, ref string value)
        {
            bool result = false;
            if (bSaveAndValidate)
            {
                ListViewItem item = listView1.Items[Row];
                if (null != item)
                {
                    if (0 == iSubItem)
                    {
                        value = item.Text;
                        result = true;
                    }
                    else
                    {
                        value = item.SubItems[iSubItem].Text;
                        result = true;
                    }
                }
            }
            else
            {
                if (0 == iSubItem && listView1.Items.Count <= Row)
                {
                    listView1.Items.Add(value, 4);
                    result = true;
                }
                else
                {
                    ListViewItem item = listView1.Items[Row];
                    if (null != item)
                    {
                        if (item.SubItems.Count <= iSubItem)
                        {
                            ListViewItem.ListViewSubItem subitem = item.SubItems.Add(value);
                            if (null != subitem)
                            {
                                listView1.SetItemImage(Row, iSubItem, 4);
                                result = true;
                            }
                        }
                        else
                        {
                            item.SubItems[iSubItem].Text = value;
                            result = true;
                        }
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// ����/�u���C�x���g�n���h���[
        /// </summary>
        /// <param name="code"></param>
        /// <param name="e"></param>
        private void findReplaceHandler(int code, EventArgs e)
        {
            switch (code)
            {
                case 1:
                    repeatToolStripMenuItem_Click(null, e);
                    break;
                case 2:
                    repeatToolStripMenuItem_Click(null, e);
                    break;
                case 3:
                    break;
                default:
                    m_find1 = null;
                    m_replace1 = null;
                    break;
            }
        }
    }
}