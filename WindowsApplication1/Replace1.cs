using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace WindowsApplication1
{
    public partial class Replace1 : Form
    {
        public delegate void FindReplaceHandler(int code, EventArgs e);
        public FindReplaceHandler m_eventHandler;
        public Replace1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            bool bEnabled = 0 < textBox1.Text.Length;
            button1.Enabled = bEnabled;
            button2.Enabled = bEnabled;
            button3.Enabled = bEnabled;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (null != m_eventHandler)
            {
                m_eventHandler(1, new EventArgs());
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (null != m_eventHandler)
            {
                m_eventHandler(2,new EventArgs());
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (null != m_eventHandler)
            {
                m_eventHandler(3, new EventArgs());
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Replace1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (null != m_eventHandler)
            {
                m_eventHandler(0, new EventArgs());
            }
        }
    }
}