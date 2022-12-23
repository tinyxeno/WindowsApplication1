using System.Collections;
using System.Windows.Forms;

namespace WindowsApplication1
{
    public class ListViewColumnSorter : IComparer
    {
        /// <summary>
        /// ���בւ����
        /// </summary>
        private int m_column;
        /// <summary>
        /// ���בւ������
        /// </summary>
        private SortOrder m_direction;

        /// <summary>
        /// �R���X�g���N�^
        /// </summary>
        public ListViewColumnSorter()
        {
            Clear();
        }

        /// <summary>
        /// ���בւ����ʒu
        /// </summary>
        public int SortColumn
        {
            set
            {
                m_column = value;
            }
            get
            {
                return m_column;
            }
        }

        /// <summary>
        /// ���בւ������
        /// </summary>
        public SortOrder Order
        {
            set
            {
                m_direction = value;
            }
            get
            {
                return m_direction;
            }
        }

        /// <summary>
        /// ���בւ������̃N���A
        /// </summary>
        public void Clear()
        {
            m_column = 0;
            m_direction = SortOrder.None;
        }

        /// <summary>
        /// ���X�g�r���[�̃A�C�e���E�I�u�W�F�N�g�̑召�֌W�𔻒肷��
        /// </summary>
        /// <param name="arg1">���X�g�r���[�̃A�C�e��1</param>
        /// <param name="arg2">���X�g�r���[�̃A�C�e��2</param>
        /// <returns>�e�A�C�e���̔�r���ʁB�召�֌W��Ԃ��B�������ꍇ��0�B</returns>
        public int Compare(object item1, object item2)
        {
            int result = 0;
            switch (m_direction)
            {
                case SortOrder.None:
                    break;
                default:
                    if (null != item1 && null != item2)
                    {
                        ListViewItem
                            enter = (ListViewItem)item1,
                            leave = (ListViewItem)item2;
                        string value = enter.SubItems[m_column].Text;
                        result = value.CompareTo(leave.SubItems[m_column].Text);
                        switch (m_direction)
                        {
                            case SortOrder.Descending:
                                result = -result;
                                break;
                        }
                    }
                    break;
            }
            return result;
        }
    }
}
