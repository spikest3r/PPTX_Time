using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace PPTX_Time
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private List<PPT> presentations;
        int totalMinutes;

        private void UpdateList()
        {
            listBox1.Items.Clear();
            foreach(var item in presentations)
            {
                listBox1.Items.Add(item.name);
            }
        }

        private void CalculateTotalTime()
        {
            float time = 0.0f;
            string scale = "minutes";
            FormatTime(totalMinutes, out time, out scale, checkBox1.Checked);
            label1.Text = string.Format("Total editing time:\n {0} {1}", time, scale);
        }

        private int ProcessFile(string file)
        {
            int minutesValue = 0;
            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(file))
                {
                    ZipArchiveEntry coreEntry = archive.GetEntry("docProps/app.xml");
                    ZipArchiveEntry thumbnailEntry = archive.GetEntry("docProps/thumbnail.jpeg");
                    if (coreEntry != null)
                    {
                        using (Stream stream = coreEntry.Open())
                        {
                            PPT ppt = new PPT();

                            ppt.name = Path.GetFileNameWithoutExtension(file);

                            XmlDocument xmlDoc = new XmlDocument();
                            xmlDoc.Load(stream);

                            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                            nsmgr.AddNamespace("ep", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");

                            XmlNode totalTimeNode = xmlDoc.SelectSingleNode("//ep:TotalTime", nsmgr);
                            if (totalTimeNode != null && int.TryParse(totalTimeNode.InnerText, out int minutes))
                            {
                                minutesValue = minutes;
                                ppt.rawTime = minutes;
                            }
                            else
                            {
                                ppt.rawTime = -1;
                            }

                            if (thumbnailEntry != null)
                            {
                                using (Stream entryStream = thumbnailEntry.Open())
                                using (MemoryStream ms = new MemoryStream())
                                {
                                    entryStream.CopyTo(ms);
                                    ms.Position = 0;
                                    try
                                    {
                                        using (var tmp = Image.FromStream(ms))
                                        {
                                            ppt.thumbnail = new Bitmap(tmp); // Clean copy
                                        }
                                    }
                                    catch { ppt.thumbnail = null; }
                                }
                            } else
                            {
                                ppt.thumbnail = null;
                            }

                            presentations.Add(ppt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Failed to file\n{0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return minutesValue;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                totalMinutes = 0;

                pictureBox1.Image = null;

                foreach (var item in presentations)
                {
                    if (item.thumbnail == null) continue;
                    item.thumbnail.Dispose();
                }

                presentations.Clear();

                GC.Collect();
                GC.WaitForPendingFinalizers();

                string folderPath = folderBrowserDialog1.SelectedPath;
                foreach (string file in Directory.GetFiles(folderPath, "*.pptx"))
                {
                    totalMinutes += ProcessFile(file);
                }

                UpdateList();

                CalculateTotalTime();
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        
        private void FormatTime(int inTime, out float time, out string scale, bool forceHours = false)
        {
            string[] scales = { "minutes", "hours", "days" };
            int scaleIndex = 0;
            float output = inTime;
            if((inTime >= 60 && scaleIndex == 0) || forceHours)
            {
                scaleIndex++;
                output = inTime / 60.0f; // minutes to hours
            }
            if (output >= 24 && scaleIndex == 1 && !forceHours)
            {
                scaleIndex++;
                output /= 24.0f; // hours to days
            }
            time = output;
            scale = scales[scaleIndex];
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int newIndex = listBox1.SelectedIndex;
            if (newIndex < 0)
            {
                label2.Text = "Name";
                label3.Text = "Editing time";
                button3.Enabled = false;
                pictureBox1.Image = null;
                return;
            }

            var presentation = presentations[newIndex];
            
            string scale = "minutes";
            float time = 0.0f;
            FormatTime(presentation.rawTime, out time, out scale, checkBox1.Checked);

            label2.Text = string.Format("Name: {0}", presentation.name);
            label3.Text = string.Format("Editing time: {0} {1}", time, scale);

            button3.Enabled = true;

            pictureBox1.Image = presentation.thumbnail;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            presentations = new List<PPT>();
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog1.FileName;
                int minutes = ProcessFile(path);

                totalMinutes += minutes;
                UpdateList();

                CalculateTotalTime();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var presentation = presentations[listBox1.SelectedIndex];
            presentation.thumbnail?.Dispose();
            totalMinutes -= presentation.rawTime;
            presentations.RemoveAt(listBox1.SelectedIndex);
            listBox1.ClearSelected();
            UpdateList();
            CalculateTotalTime();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            CalculateTotalTime(); // recalculate to fit new params
        }

        private bool _closing = false;

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_closing) return;
            _closing = true;

            var img = pictureBox1.Image;
            pictureBox1.Image = null;
            img?.Dispose();

            base.OnFormClosing(e);
        }
    }

    public class PPT
    {
        public string name;
        public int rawTime; // minutes
        public Image thumbnail;
    }
}
