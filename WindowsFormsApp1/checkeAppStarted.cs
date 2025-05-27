using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    class checkeAppStarted
    {
        public void IskAppStart()
        {
            bool Is_createNew = false;
            Mutex mu = null;
            string mutexName = Process.GetCurrentProcess().MainModule.FileName.Replace(Path.DirectorySeparatorChar, '_');
            //MessageBox.Show(mutexName);

            mu = new Mutex(true, "Global\\" + mutexName, out Is_createNew);
            if (!Is_createNew)
            {
                MessageBox.Show("程式已開啟!");
                //this.Close();
                Application.Exit();
            }
        }
    }
}
