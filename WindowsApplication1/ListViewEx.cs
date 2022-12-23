using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace WindowsApplication1
{
    public class ListViewEx : ListView
    {
        /// <summary>
        /// ��`�\����
        /// </summary>
        private struct RECT
        {
            public Int32 nLeft;
            public Int32 nTop;
            public Int32 nRight;
            public Int32 nBottom;
        };

        /// <summary>
        /// �A�C�e���\���́B�T�u�A�C�e���̃C���[�W���̎擾�Ɛݒ�p
        /// </summary>
        private struct LVITEM
        {
            public UInt32 mask;
            public Int32 iItem;
            public Int32 iSubItem;
            public UInt32 state;
            public UInt32 stateMask;
            public string pszText;
            public Int32 cchTextMax;
            public Int32 iImage;
            public IntPtr lParam;
            public Int32 iIndent;
            public Int32 iGroupId;
            public UInt32 cColumns;
            public UInt32 puColumns;
            public IntPtr piColFmt;
            public Int32 iGroup;
        };

        /// <summary>
        /// ��ݒ�\���́B�J�����w�b�_�̃\�[�g�C���W�P�[�^�̏����ŗ��p
        /// </summary>
        private struct LVCOLUMN
        {
            public UInt32 mask;
            public Int32 fmt;
            public Int32 cx;
            public string pszText;
            public Int32 cchTextMax;
            public Int32 iSubItem;
            public Int32 iImage;
            public Int32 iOrder;
            public Int32 cxMin;
            public Int32 cxDefault;
            public Int32 cxIdeal;
        };

        /// <summary>
        /// ����ݒ�_�C�A���O�p�\����
        /// </summary>
        private struct PRINTDLG
        {
            public UInt32 lStructSize;
            public IntPtr hwndOwner;
            public IntPtr hDevMode;
            public IntPtr hDevNames;
            public IntPtr hDC;
            public UInt32 Flags;
            public UInt16 nFromPage;
            public UInt16 nToPage;
            public UInt16 nMinPage;
            public UInt16 nMaxPage;
            public UInt16 nCopies;
            public IntPtr hInstance;
            public IntPtr lCustData;
            public IntPtr lpfnPrintHook;
            public IntPtr lpfnSetupHook;
            string lpPrintTemplateName;
            string lpSetupTemplateName;
            public IntPtr hPrintTemplate;
            public IntPtr hSetupTemplate;
        };


        private const Int32 LVIF_IMAGE  = 2;
        private const Int32 LVCF_FMT    = 1;
        private const Int32 LVIR_BOUNDS = 0;
        private const Int32 LVIR_ICON   = 1;
        private const Int32 LVIR_LABEL  = 2;
        private const Int32 LVIR_SELECTBOUNDS   = 3;
        private const Int32 LVS_EX_SUBITEMIMAGES    = 2;

        private const Int32 HDF_SORTUP      = 0x0400;
        private const Int32 HDF_SORTDOWN    = 0x0200;

        private const Int32 PD_PRINTSETUP   = 0x0040;

        private const Int32 CF_UNICODETEXT  = 13;

        private const Int32 WM_PAINT        = 0xF;

        private const Int32 EM_GETSEL       = 0xB0;
        private const Int32 EM_SETSEL       = 0xB1;
        private const Int32 EM_CANUNDO      = 0xC6;
        private const Int32 EM_UNDO         = 0xC7;
        private const Int32 WM_CUT          = 0x300;
        private const Int32 WM_COPY         = 0x301;
        private const Int32 WM_PASTE        = 0x302;

        private const Int32 LVM_FIRST                       = 0x1000;
        private const Int32 LVM_GETEDITCONTROL              = LVM_FIRST + 24;
        private const Int32 LVM_SETEXTENDEDLISTVIEWSTYLE    = LVM_FIRST + 54;
        private const Int32 LVM_GETEXTENDEDLISTVIEWSTYLE    = LVM_FIRST + 55;
        private const Int32 LVM_GETSUBITEMRECT              = LVM_FIRST + 56;
        private const Int32 LVM_GETITEM                     = LVM_FIRST + 75;
        private const Int32 LVM_SETITEM                     = LVM_FIRST + 76;
        private const Int32 LVM_GETCOLUMNW                  = LVM_FIRST + 95;
        private const Int32 LVM_SETCOLUMNW                  = LVM_FIRST + 96;

        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern IntPtr SendMessage(IntPtr hWnd, Int32 Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern IntPtr SendMessageRef(IntPtr hWnd, Int32 Msg, ref Int32 wParam, ref Int32 lParam);

        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern IntPtr SendMessageRECT(IntPtr hWnd, Int32 Msg, Int32 wParam, ref RECT lParam);

        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern IntPtr SendMessageLVITEM(IntPtr hWnd, Int32 Msg, IntPtr wParam, ref LVITEM lParam);

        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern IntPtr SendMessageLVCOLUMN(IntPtr hWnd, Int32 Msg, IntPtr wParam, ref LVCOLUMN lParam);

        [DllImport("User32.dll")]
        private static extern IntPtr MoveWindow(IntPtr hWnd, Int32 x, Int32 y, Int32 w, Int32 h, bool redraw);

        [DllImport("User32.dll")]
        private static extern Int32 GetWindowTextLength(IntPtr hWnd);
        [DllImport("User32.dll")]
        private static extern IntPtr SetWindowText(IntPtr hWnd, string text);

        [DllImport("User32.dll")]
        private static extern IntPtr IsClipboardFormatAvailable(Int32 format);

        [DllImport("User32.dll")]
        private static extern IntPtr GetFocus();

        [DllImport("Comdlg32.dll")]
        private static extern Int32 PrintDlg(ref�@PRINTDLG printDlg);

        /// <summary>
        /// ���ڂ��Ă���s��ʒu
        /// </summary>
        private int m_iItem, m_iSubItem;

        /// <summary>
        /// �R���X�g���N�^
        /// </summary>
        public ListViewEx()
        {
            m_iItem = m_iSubItem = 0;
        }

        /// <summary>
        /// �t���[�e�B���O�E�G�f�B�^�̃n���h���̎擾
        /// </summary>
        private IntPtr FloatEdit
        {
            get
            {
                return SendMessage(Handle, LVM_GETEDITCONTROL, IntPtr.Zero, IntPtr.Zero);
            }
        }

        /// <summary>
        /// �s�ʒu
        /// </summary>
        public int Row
        {
            get
            {
                return m_iItem;
            }
            set
            {
                m_iItem = value;
            }
        }

        /// <summary>
        /// ��ʒu
        /// </summary>
        public int Col
        {
            get
            {
                return m_iSubItem;
            }
            set
            {
                m_iSubItem = value;
            }
        }

        /// <summary>
        /// �y�ݒ�̂݁z�t���[�e�B���O�G�f�B�^�֕������ݒ肷��
        /// </summary>
        public string LabelString
        {
            set
            {
                IntPtr hEdit = FloatEdit;
                if (IntPtr.Zero != hEdit)
                {
                    SetWindowText(hEdit, value);
                }
            }
        }

        /// <summary>
        /// �؂��菈��
        /// </summary>
        public bool Cut
        {
            /// �؂���ł��邩������Ԃ�
            get
            {
                bool result = false;
                IntPtr hEdit = FloatEdit;
                if (IntPtr.Zero != hEdit)
                {
                    Int32 enter = -1, leave = -1;
                    SendMessageRef(hEdit, EM_GETSEL, ref enter, ref leave);
                    result = -1 != enter && -1 != leave && enter != leave;
                }
                return result;
            }
            /// �؂��菈�������s����y�^��ݒ肵���ꍇ�̂݁z
            set
            {
                if (value)
                {
                    IntPtr hEdit = FloatEdit;
                    if (IntPtr.Zero != hEdit)
                    {
                        SendMessage(hEdit, WM_CUT, IntPtr.Zero, IntPtr.Zero);
                    }
                }
            }
        }

        /// <summary>
        /// �R�s�[����
        /// </summary>
        public bool Copy
        {
            /// �R�s�[�������ł��邩�ǂ����𔻒肷��
            get
            {
                return Cut;
            }
            /// �R�s�[���������s����y�^��ݒ肵���ꍇ�̂݁z
            set
            {
                if (value)
                {
                    IntPtr hEdit = FloatEdit;
                    if (IntPtr.Zero != hEdit)
                    {
                        SendMessage(hEdit, WM_COPY, IntPtr.Zero, IntPtr.Zero);
                    }
                }
            }
        }

        /// <summary>
        /// �\��t������
        /// </summary>
        public bool Paste
        {
            /// �\��t���������ł��邩�ǂ����𔻒肷��
            get
            {
                bool result = false;
                IntPtr hEdit = FloatEdit;
                if (IntPtr.Zero != hEdit)
                {
                    result = IntPtr.Zero != IsClipboardFormatAvailable(CF_UNICODETEXT);
                }
                return result;
            }
            /// �\��t�����������s����y�^��ݒ肵���ꍇ�̂݁z
            set
            {
                if (value)
                {
                    IntPtr hEdit = FloatEdit;
                    if (IntPtr.Zero != hEdit)
                    {
                        SendMessage(hEdit, WM_PASTE, IntPtr.Zero, IntPtr.Zero);
                    }
                }
            }
        }

        /// <summary>
        /// ���ɖ߂�����
        /// </summary>
        public bool Undo
        {
            /// ���ɖ߂��������ł��邩�ǂ����𔻒肷��
            get
            {
                IntPtr hEdit = FloatEdit;
                return IntPtr.Zero != hEdit &&
                    IntPtr.Zero != SendMessage(hEdit, EM_CANUNDO, IntPtr.Zero, IntPtr.Zero);
            }
            /// ���ɖ߂����������s����y�^��ݒ肵���ꍇ�̂݁z
            set
            {
                if (value)
                {
                    IntPtr hEdit = FloatEdit;
                    if (IntPtr.Zero != hEdit)
                    {
                        SendMessage(hEdit, EM_UNDO, IntPtr.Zero, IntPtr.Zero);
                    }
                }
            }
        }

        /// <summary>
        /// ���ׂđI������
        /// </summary>
        public bool SelectAll
        {
            /// ���ׂđI���������ł��邩�ǂ����𔻒肷��
            get
            {
                return IntPtr.Zero != FloatEdit;
            }
            /// ���ׂđI�����������s����y�^��ݒ肵���ꍇ�̂݁z
            set
            {
                if (value)
                {
                    IntPtr hEdit = FloatEdit;
                    if (IntPtr.Zero != hEdit)
                    {
                        int length = GetWindowTextLength(hEdit);
                        SendMessage(hEdit, EM_SETSEL, IntPtr.Zero, (IntPtr)length);
                    }
                }
            }
        }

        /// <summary>
        /// ���X�g�r���[�փt�H�[�J�X���������Ă��邩�ǂ����𔻒肷��
        /// </summary>
        public bool IsFocusListView
        {
            get
            {
                bool result = false;
                IntPtr value = GetFocus();
                if (IntPtr.Zero != value)
                {
                    result = Handle == value;
                }
                return result;
            }
        }

        /// <summary>
        /// �T�u�A�C�e���փC���[�W��ݒ肷�邩�ǂ�����ݒ�/�擾����
        /// </summary>
        public bool SubItemImages
        {
            get
            {
                IntPtr data = SendMessage(Handle, LVM_GETEXTENDEDLISTVIEWSTYLE, IntPtr.Zero, IntPtr.Zero);
                return 0 != (LVS_EX_SUBITEMIMAGES & data.ToInt32());
            }
            set
            {
                Int32 mask = LVS_EX_SUBITEMIMAGES;
                IntPtr data = SendMessage(Handle, LVM_GETEXTENDEDLISTVIEWSTYLE, IntPtr.Zero, IntPtr.Zero);
                Int32 style = data.ToInt32();
                style &= ~mask;
                if (value)
                {
                    style |= mask;
                }
                SendMessage(Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, IntPtr.Zero, (IntPtr)style);
            }
        }

        /// <summary>
        /// �T�u�A�C�e���֐ݒ肳��Ă���C���[�W�ԍ����擾����
        /// </summary>
        /// <param name="iItem">�s�ʒu</param>
        /// <param name="iSubItem">��ʒu</param>
        /// <returns>�T�u�A�C�e���֐ݒ肳��Ă���C���[�W�̃C���f�b�N�X�l</returns>
        public Int32 GetItemImage(Int32 iItem, Int32 iSubItem)
        {
            Int32 result = 0;
            LVITEM item     = new LVITEM();
            item.iItem      = iItem;
            item.iSubItem   = iSubItem;
            item.mask       = LVIF_IMAGE;
            if (IntPtr.Zero != SendMessageLVITEM(Handle, LVM_GETITEM, IntPtr.Zero, ref item))
            {
                result = item.iImage;
            }
            return result;
        }

        /// <summary>
        /// �T�u�A�C�e���փC���[�W�ԍ���ݒ肷��
        /// </summary>
        /// <param name="iItem">�s�ʒu</param>
        /// <param name="iSubItem">��ʒu</param>
        /// <param name="value">�C���[�W�̃C���f�b�N�X�l</param>
        public void SetItemImage(Int32 iItem, Int32 iSubItem, Int32 value)
        {
            LVITEM item     = new LVITEM();
            item.iItem      = iItem;
            item.iSubItem   = iSubItem;
            item.mask       = LVIF_IMAGE;
            item.iImage     = value;
            SendMessageLVITEM(Handle, LVM_SETITEM, IntPtr.Zero, ref item);
        }

        /// <summary>
        /// �w��ʒu�̋�`�̈���擾����
        /// </summary>
        /// <param name="iItem">�s�ʒu</param>
        /// <param name="iSubItem">��ʒu</param>
        /// <returns>��`�̈�</returns>
        public Rectangle GetItemBound(Int32 iItem, Int32 iSubItem)
        {
            Rectangle result = new Rectangle();
            RECT rc = new RECT();
            rc.nLeft    = LVIR_LABEL;
            rc.nTop     = iSubItem;
            rc.nRight   = 0;
            rc.nBottom  = 0;
            if (IntPtr.Zero != SendMessageRECT(Handle, LVM_GETSUBITEMRECT, iItem, ref rc))
            {
                result.X        = rc.nLeft;
                result.Y        = rc.nTop;
                result.Width    = rc.nRight - rc.nLeft;
                result.Height   = rc.nBottom - rc.nTop;
            }
            return result;
        }

        /// <summary>
        /// �J�����w�b�_�[�̃\�[�g�E�C���W�P�[�^�[�̏�Ԃ�Ԃ�
        /// </summary>
        /// <param name="index">��ʒu</param>
        /// <returns>�ݒ肳��Ă���\�[�g��Ԃ�Ԃ�</returns>
        public SortOrder GetSortOrder(int index)
        {
            SortOrder result = SortOrder.None;
            LVCOLUMN source = new LVCOLUMN();
            source.mask = LVCF_FMT;
            if (IntPtr.Zero != SendMessageLVCOLUMN(Handle, LVM_GETCOLUMNW, (IntPtr)index, ref�@source))
            {
                switch ((HDF_SORTDOWN | HDF_SORTUP) & source.fmt)
                {
                    case HDF_SORTUP:
                        result = SortOrder.Ascending;
                        break;
                    case HDF_SORTDOWN:
                        result = SortOrder.Descending;
                        break;
                }
            }
            return result;
        }
        /// <summary>
        /// �J�����w�b�_�[�̃\�[�g�E�C���W�P�[�^�[��ݒ肷��
        /// </summary>
        /// <param name="index">��ʒu</param>
        /// <param name="value">�\�[�g�C���W�P�[�^�[�̏��</param>
        public void SetSortOrder(int index, SortOrder value)
        {
            LVCOLUMN source = new LVCOLUMN();
            source.mask = LVCF_FMT;
            if (IntPtr.Zero != SendMessageLVCOLUMN(Handle, LVM_GETCOLUMNW, (IntPtr)index, ref�@source))
            {
                source.fmt &= ~(HDF_SORTDOWN | HDF_SORTUP);
                switch (value)
                {
                    case SortOrder.Ascending:
                        source.fmt |= HDF_SORTUP;
                        break;
                    case SortOrder.Descending:
                        source.fmt |= HDF_SORTDOWN;
                        break;
                }
                SendMessageLVCOLUMN(Handle, LVM_SETCOLUMNW, (IntPtr)index, ref�@source);
            }
        }

        /// <summary>
        /// ����ݒ�_�C�A���O��\������
        /// </summary>
        /// <returns>��������</returns>
        public DialogResult ShowPrintSetupDlg()
        {
            DialogResult result = DialogResult.Cancel;
            PRINTDLG printDlg = new PRINTDLG();
            printDlg.lStructSize = 120;
            printDlg.hwndOwner = Handle;
            printDlg.Flags = PD_PRINTSETUP;
            if (0 != PrintDlg(ref printDlg))
            {
                result = DialogResult.OK;
            }
            return result;
        }

        /// <summary>
        /// ���b�Z�[�W����
        /// </summary>
        /// <param name="m">���b�Z�[�W�\����</param>
        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_PAINT:
                    if (true)
                    {
                        IntPtr hEdit = FloatEdit;
                        if (IntPtr.Zero != hEdit)
                        {
                            RECT rc     = new RECT();
                            rc.nLeft    = LVIR_LABEL;
                            rc.nTop     = m_iSubItem;
                            rc.nRight   = 0;
                            rc.nBottom  = 0;
                            if (IntPtr.Zero != SendMessageRECT(Handle, LVM_GETSUBITEMRECT, m_iItem, ref rc))
                            {
                                MoveWindow(hEdit, rc.nLeft, rc.nTop, rc.nRight - rc.nLeft, rc.nBottom - rc.nTop, true);
                            }
                        }
                    }
                    break;
            }
            base.WndProc(ref m);
        }
    }
}
