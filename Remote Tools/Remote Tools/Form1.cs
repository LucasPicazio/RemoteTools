using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Management;
using System.Data;
using System.IO;
using Microsoft.Win32;
using System.Security.Principal;
using System.Data.OleDb;
using System.DirectoryServices.AccountManagement;
using System.Diagnostics;

namespace WindowsFormsApp6
{
    public partial class Form1 : Form
    {
       
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();


        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        // Controles e chamadas de metódos
        // atualizaPlanilha() é chamado em todos as opçoes para que seja verificado se a planilha esta atualizada de acordo com a data
        //e se não estiver verificar sem reiniciar o programa
        {
            if (radioButton1.Checked == true)
            {
                IQP(GetBRD());
                atualizaPlanilha();
            }
            if (radioButton2.Checked == true)
            {
                IQP(GetBRD());
                perfilOut(GetBRD());
                atualizaPlanilha();
            }
            if (radioButton3.Checked == true)
            {
                arrumaTeclado(GetBRD());
                atualizaPlanilha();
            }
            if (radioButton4.Checked == true)
            {
                Offer(GetBRD());
                atualizaPlanilha();
            }
            if (radioButton5.Checked == true)
            {
                switchOn(GetBRD());
                atualizaPlanilha();
            }
            if (radioButton6.Checked == true)
            {
                atualizaPlanilha();
                switchOFF(GetBRD());
            }
            if (radioButton7.Checked == true)
            {

                atualizaPlanilha();
                MapIeJ(GetBRD());
            }
            if (radioButton8.Checked == true)
            {
                atualizaPlanilha();
                mapPrinter(GetBRD(), textBox2.Text);
            }
        }

        private void mapPrinter(string BRD, string Nprinter)
        // Metodo para mapear impressoras rodando um arquvio .bat na instancia do usuario. Usando WMI
        {
           string sPrinterName = @"\\csbrprtsrv.la.hedani.net\P_SP00" + Nprinter;
            var connectoptions = new ConnectionOptions();
            connectoptions.Impersonation = ImpersonationLevel.Default;
            connectoptions.EnablePrivileges = true;

            ManagementScope scope = new ManagementScope(@"\\" + BRD + @"\root\cimv2", connectoptions);
            scope.Connect();

            ManagementClass oPrinterClass = new ManagementClass (new ManagementPath("Win32_Printer"),null);
            ManagementBaseObject oInputParameters = oPrinterClass.GetMethodParameters("AddPrinterConnection");
           
            oInputParameters.SetPropertyValue("Name", sPrinterName);

            oPrinterClass.InvokeMethod("AddPrinterConnection", oInputParameters,null);

        }

        private void MapIeJ(string BRD)
        // Metodo para mapear pastas compartilhadas na instancia do usuario. No .bat tem dois comando netuse. Usando WMI para criar um processo;
        {
            var processToRun = new[] { "notepad.exe" };
            var connectoptions = new ConnectionOptions();
            ManagementScope scope = new ManagementScope(@"\\" + BRD + @"\root\cimv2", connectoptions);
            var wmiProcess = new ManagementClass(scope, new ManagementPath("Win32_Process"), new ObjectGetOptions());
            wmiProcess.InvokeMethod("Create", processToRun);
            
        }
        
        private void switchOFF(string BRD)
        // Método para trocar chave do registro que tira do computador a opção de trocar de usuario na tela de Login
        {
            RegistryKey k = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, BRD)
                .OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", true);

            k.SetValue("HideFastUserSwitching", 1, RegistryValueKind.DWord);
            ;
        }

        private void switchOn(string BRD)
        // Método para trocar chave do registro que poe no computador a opção de trocar de usuario na tela de Login
        {
            RegistryKey k = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, BRD)
                            .OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", true);

            k.SetValue("HideFastUserSwitching", 0, RegistryValueKind.DWord);
        }

        private void Offer(string BRD)
        // Método para iniciar o processo de remote assistance com argumentos de Offer e nome da Maquina
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();

            System.Diagnostics.ProcessStartInfo startInfo =
            new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Maximized;
            startInfo.FileName = "msra.exe";


            startInfo.Arguments = "/offerRA " + BRD;

            process.StartInfo = startInfo;
            process.Start();

        }
        private void arrumaTeclado(string BRD)
        //Metodo que troca chave de registros do teclado na instancia do usuario e deixa no padrão funcional.
        {

            RegistryKey k = RegistryKey.OpenRemoteBaseKey(RegistryHive.Users, BRD);
            RegistryKey ksubstitute = RegistryKey.OpenRemoteBaseKey(RegistryHive.Users, BRD);

            foreach (string sid in k.GetSubKeyNames())
            {
                try
                {
                    SecurityIdentifier s = new SecurityIdentifier(sid);
                    string x = s.Translate(typeof(NTAccount)).ToString();
                    if (x.EndsWith(textBox1.Text.Substring(0, textBox1.Text.IndexOf(" "))))
                    {
                        k = k.OpenSubKey(sid).OpenSubKey("Keyboard Layout").OpenSubKey("Preload", true);
                        ksubstitute = k.OpenSubKey(sid).OpenSubKey("Keyboard Layout").OpenSubKey("Substitutes", true);


                        foreach (string value in k.GetValueNames())
                        {
                            k.DeleteValue(value);
                        }
                        foreach (string value in ksubstitute.GetValueNames())
                        {
                            ksubstitute.DeleteValue(value);
                        }
                        k.SetValue("1", "00000409", RegistryValueKind.String);
                        ksubstitute.SetValue("0000540a", "0000080a", RegistryValueKind.String);
                        k.Close();
                        ksubstitute.Close();
                        break;
                    }
                }
                catch (Exception t)
                {

                }
            }






        }

        private void perfilOut(string BRD)
        // Método para mudar nome do arquivo dentro da maquina do usuario para resetar perfil do Outlook
        {
            DirectoryInfo hdDirectoryInWhichToSearch = new DirectoryInfo(@"\\" + BRD + @"\C$\Users\" + textBox1.Text.Substring(0, textBox1.Text.IndexOf('-')).TrimEnd() + @"\AppData\Local\Microsoft\Outlook");
            FileInfo[] filesInDir = hdDirectoryInWhichToSearch.GetFiles("*.ost");

            foreach (FileInfo foundFile in filesInDir)
            {
                string fullName = foundFile.FullName;
                MessageBox.Show(fullName);
                Random rnd = new Random();
                String number = rnd.Next().ToString();
                File.Move(fullName, fullName + number + ".old");
            }

        }

        private string GetBRD()
        // Método usa nome do usuario fornecido na textbox e procura em datatable para conseguir nome da maquina dele
        {
            if (textBox1.Text.StartsWith("BRD", true, null) || textBox1.Text.StartsWith("BRM", true, null)) return textBox1.Text;
            var drows = dt.Select("[Who is] like '%" + textBox1.Text.Substring(0,textBox1.Text.IndexOf(" ")) + "%'");
            if(drows.Count() > 1)
            {
                Form2 f2 = new Form2();
                List<CheckBox> list = new List<CheckBox>();
                list.Add(f2.checkBox1);
                list.Add(f2.checkBox2);
                list.Add(f2.checkBox3);
                list.Add(f2.checkBox4);

                int i = 0;
                foreach(CheckBox cb in list)
                {
                    
                    cb.Visible = false;
                    i++;
                }

                for(int t = 0; t<drows.Count(); t++)
                {
                    list[t].Text = drows[t].ItemArray[0].ToString();
                    list[t].Visible = true;
                }


                f2.ShowDialog();
                return f2.BRX;
            }
            return (string)drows[0].ItemArray[0];
        }

        private void IQP(string ipAddress)
        //metódo mata processo da lista abaixo via WMI
        {
            var processName = @"SI.IQProtector64.exe";
            var processName2 = @"SI.WebProxy.exe";
            var processName3 = @"OUTLOOK.exe";
            var processName4 = @"Lync.exe";
            var processName5 = @"UcMapi.exe";

            var connectoptions = new ConnectionOptions();

            ManagementScope scope = new ManagementScope(@"\\" + ipAddress + @"\root\cimv2", connectoptions);

            var query = new SelectQuery("select * from Win32_process where name = '" + processName +
                "'OR name ='" + processName2 + "'OR name ='" + processName3 + "'OR name ='" + processName3 + "'OR name ='" + processName4 + "'OR name ='" + processName5 + "'");

            using (var searcher = new ManagementObjectSearcher(scope, query))
            {
                foreach (ManagementObject process in searcher.Get())
                {
                    try
                    {
                        process.InvokeMethod("Terminate", null);
                    }
                    catch (Exception e)
                    {

                    }

                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        // Método para ser carregado no carregamento do Form
        {

            label1.Visible = false;
            textBox2.Visible = false;

            atualizaPlanilha();
            radioButton8.Visible = false;
            radioButton7.Visible = false;


        }



        private void atualizaPlanilha()
        // Método transforma planilha csv e xls em datatable. Usa da Xls para criar uma planilha CSV com o login e o nome completo do usuario
        // Após passar para datatable, pega uma coluna, transforma em array de String e usa de Fonte para o AutoCompleteCustomSource do textbox
        // OleDB para planilha xls e StreamReader para arquivo CSV

        {
            // os paths são feitos com as datas para facilitar a diferenciação de planilhas no acesso
            var path = @"\\csao11p20011d\rwapps\CSHG\CSHGDsl\SUPORTE\RemoteTools\rs" + DateTime.Now.ToString("_yyyy_MM_dd") + ".xls";
            var path2 = @"\\csao11p20011d\rwapps\CSHG\CSHGDsl\SUPORTE\RemoteTools\rs" + DateTime.Now.ToString("_yyyy_MM_dd") + ".csv";
            //------------------------------------------------------- csv to datatable
            if (File.Exists(path2))
            {

                if (dt.Rows.Count != 0)
                {

                    dt.Clear();
                    

                }
                StreamReader sr = new StreamReader(path2);
                string[] headers = sr.ReadLine().Split('\t');
                if (dt.Columns.Count == 0)
                {
                    foreach (string header in headers)
                    {
                        dt.Columns.Add(header);
                    }
                }
                
                while (!sr.EndOfStream)
                {

                    string[] rows = sr.ReadLine().Split('\t');
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        dr[i] = rows[i];
                    }
                    dt.Rows.Add(dr);
                }

            }
            //------------------------------------------------------- csv to datatable

            //------------------------------------------------------- xls to datatale e to csv
            if (!File.Exists(path2))
            {
                string excelConnStr = string.Empty;
                OleDbCommand excelCommand = new OleDbCommand();
                OleDbDataAdapter excelDataAdapter = new OleDbDataAdapter();


                excelConnStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + path + ";Extended Properties=Excel 12.0";


                OleDbConnection excelConn = new OleDbConnection(excelConnStr);
                excelConn.Open();

                excelCommand = new OleDbCommand("SELECT * FROM [Sheet1$]", excelConn);
                excelDataAdapter.SelectCommand = excelCommand;
                excelDataAdapter.Fill(dt);



                //-------------------------------------------------------xls to datatale

                //-------------------------------------------------------Replace de datatable

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i][2] = dt.Rows[i][2].ToString().Replace("GBLADHEDANI\\", "");

                    using (PrincipalContext context = new PrincipalContext(ContextType.Domain, "gbl.ad.hedani.net", "OU=BR,OU=CS,DC=gbl,DC=ad,DC=hedani,DC=net"))
                    {
                        try
                        {
                            using (UserPrincipal usr = UserPrincipal.FindByIdentity(context, dt.Rows[i][2].ToString()))
                            {
                                dt.Rows[i][2] = dt.Rows[i][2].ToString() + " - " + usr.DisplayName;
                            }
                        }
                        catch (Exception e)
                        {

                        }


                    }
                }
                excelConn.Close();

                //------------------------- dt to csv ----------------------------------
                var lines = new List<string>();

                string[] columnNames = dt.Columns.Cast<DataColumn>().
                                                  Select(column => column.ColumnName).
                                                  ToArray();

                var header = string.Join("\t", columnNames);
                lines.Add(header);

                var valueLines = dt.AsEnumerable()
                                   .Select(row => string.Join("\t", row.ItemArray));
                lines.AddRange(valueLines);

                File.WriteAllLines(path2, lines);

                //-------------------------dt to csv ----------------------------------
            }

            string[] postSource = dt.AsEnumerable().Select(r => r.Field<string>("Who is")).ToArray();
            var source = new AutoCompleteStringCollection();
            source.AddRange(postSource);
            textBox1.AutoCompleteCustomSource = source;
            textBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            

        }

        //Parte responsavel para mostrar apenas controles necessarios dependendo da função sendo utilizada
        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {

            label1.Visible = true;
            textBox2.Visible = true;
            
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            label1.Visible = false;
            textBox2.Visible = false;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            label1.Visible = false;
            textBox2.Visible = false;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            label1.Visible = false;
            textBox2.Visible = false;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            label1.Visible = false;
            textBox2.Visible = false;
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            label1.Visible = false;
            textBox2.Visible = false;
        }
        
        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            label1.Visible = false;
            textBox2.Visible = false;
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            label1.Visible = false;
            textBox2.Visible = false;
        }
    }
}
