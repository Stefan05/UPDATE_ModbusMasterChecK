using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EasyModbus;

namespace ModBusMasterChecK
{
    public partial class Form1 : Form
    {
        ModbusClient ModClient;
        bool disconnect_connect = false;
       
        public Form1()
        {
            InitializeComponent();
        }
        #region Read fct
        private void Timer_repeat_Tick(object sender, EventArgs e)
        {
            if (disconnect_connect)
            {
                if (activate_the_slave_I == true)
                {
                    ModClient.UnitIdentifier = Convert.ToByte(txt_id_S_Orange.Text);

                    if (cmb_fct_S_Orange.Text == "Read Holding")
                    {
                        
                        try {
                            int[] vals = new int[Convert.ToInt32(txt_nr_reg_S_orange.Text)];
                            vals = ModClient.ReadHoldingRegisters(int.Parse(txt_first_rgs_S_Orange.Text), Convert.ToInt32(txt_nr_reg_S_orange.Text));

                            for (int i = 0; i < Convert.ToInt32(txt_nr_reg_S_orange.Text); i++)
                            {
                                dta_grd_S_Orange.Rows[i].Cells[1].Value = vals[i];
                            }
                        }
                        catch (Exception ex)
                        {
                            activate_the_slave_I = false;
                            if (ex.Message.ToString() == "Function code not supported by master")
                            {
                                MessageBox.Show("Registrii de tip Holding registers nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                            {
                                MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            btn_switch_orange = 0;
                            dta_grd_S_Orange.Rows.Clear();
                            activ_dezactiv_slave_orange.Switched = false;
                        }
                    }

                    if (cmb_fct_S_Orange.Text == "Read Coil")
                    {
                        
                        try
                        {
                            bool[] vals = new bool[Convert.ToInt32(txt_nr_reg_S_orange.Text)];
                            vals = ModClient.ReadCoils(int.Parse(txt_first_rgs_S_Orange.Text), Convert.ToInt32(txt_nr_reg_S_orange.Text));

                            for (int i = 0; i < Convert.ToInt32(txt_nr_reg_S_orange.Text); i++)
                            {
                                dta_grd_S_Orange.Rows[i].Cells[1].Value = vals[i];
                            }
                       }
                       catch (Exception ex)
                        {
                            activate_the_slave_I = false;
                            if(ex.Message.ToString() == "Function code not supported by master")
                            {
                                MessageBox.Show("Registrii de tip coil nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }else if(ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                            {
                                MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            btn_switch_orange = 0;
                            dta_grd_S_Orange.Rows.Clear();
                            activ_dezactiv_slave_orange.Switched = false;
                        }
                    }
                }

                if (activate_the_slave_II == true)
                {
                    ModClient.UnitIdentifier = Convert.ToByte(txt_id_S_Pink.Text);

                    if (cmb_S_Pink.Text == "Read Holding")
                    {
                        
                        try {
                            int[] vals = new int[Convert.ToInt32(txt_nr_reg_S_pink.Text)];
                            vals = ModClient.ReadHoldingRegisters(int.Parse(txt_first_rgs_Pink.Text), Convert.ToInt32(txt_nr_reg_S_pink.Text));

                            for (int i = 0; i < Convert.ToInt32(txt_nr_reg_S_pink.Text); i++)
                            {
                                dta_grd_S_pink.Rows[i].Cells[1].Value = vals[i];
                            }
                        }
                        catch(Exception ex)
                        {
                            activate_the_slave_II = false;
                            if (ex.Message.ToString() == "Function code not supported by master")
                            {
                                MessageBox.Show("Registrii de tip Holding registers nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                            {
                                MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            btn_switch_pink = 0;
                            dta_grd_S_pink.Rows.Clear();
                            activ_dezactiv_slave_pink.Switched = false;
                        }
                    }

                    if (cmb_S_Pink.Text == "Read Coil")
                    {
                        
                        try {
                            bool[] vals = new bool[Convert.ToInt32(txt_nr_reg_S_pink.Text)];
                            vals = ModClient.ReadCoils(int.Parse(txt_first_rgs_Pink.Text), Convert.ToInt32(txt_nr_reg_S_pink.Text));
                            for (int i = 0; i < Convert.ToInt32(txt_nr_reg_S_pink.Text); i++)
                            {
                                dta_grd_S_pink.Rows[i].Cells[1].Value = vals[i];
                            }
                        }
                        catch (Exception ex)
                        {
                            activate_the_slave_II = false;
                            if (ex.Message.ToString() == "Function code not supported by master")
                            {
                                MessageBox.Show("Registrii de tip coil nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                            {
                                MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            btn_switch_pink = 0;
                            dta_grd_S_pink.Rows.Clear();
                            activ_dezactiv_slave_pink.Switched = false;
                        }

                       
                    }
                }

                if (activate_the_slave_III == true)
                {
                    ModClient.UnitIdentifier = Convert.ToByte(txt_id_S_Blue.Text);

                    if (cmb_S_Blue.Text == "Read Holding")
                    {
                        
                        try {
                            int[] vals = new int[Convert.ToInt32(txt_nr_reg_S_Blue.Text)];
                            vals = ModClient.ReadHoldingRegisters(int.Parse(txt_first_rgs_S_Blue.Text), Convert.ToInt32(txt_nr_reg_S_Blue.Text));

                            for (int i = 0; i < Convert.ToInt32(txt_nr_reg_S_Blue.Text); i++)
                            {
                                dta_grd_S_blue.Rows[i].Cells[1].Value = vals[i];
                            }
                        }
                        catch(Exception ex)
                        {
                            activate_the_slave_III = false;
                            if (ex.Message.ToString() == "Function code not supported by master")
                            {
                                MessageBox.Show("Registrii de tip Holding registers nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                            {
                                MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            btn_switch_blue = 0;
                            dta_grd_S_blue.Rows.Clear();
                            activ_dezactiv_slave_blue.Switched = false;
                        }
                    }

                    if (cmb_S_Blue.Text == "Read Coil")
                    {
                        
                        try {
                            bool[] vals = new bool[Convert.ToInt32(txt_nr_reg_S_Blue.Text)];
                            vals = ModClient.ReadCoils(int.Parse(txt_first_rgs_S_Blue.Text), Convert.ToInt32(txt_nr_reg_S_Blue.Text));

                            for (int i = 0; i < Convert.ToInt32(txt_nr_reg_S_Blue.Text); i++)
                            {
                                dta_grd_S_blue.Rows[i].Cells[1].Value = vals[i];
                            }
                        }
                        catch (Exception ex)
                        {
                            activate_the_slave_III = false;
                            if (ex.Message.ToString() == "Function code not supported by master")
                            {
                                MessageBox.Show("Registrii de tip coil nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                            {
                                MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            btn_switch_blue = 0;
                            dta_grd_S_blue.Rows.Clear();
                            activ_dezactiv_slave_blue.Switched = false;
                        }
                    }
                }
            }
        }
        #endregion
        bool connect = false;
        bool disconnect = true;
        #region Connect/Disconnect
        private void Btn_conectare_Click(object sender, EventArgs e)
        {
            if(connect == false) { 
                if (cmb_port.Text != "" && cmb_baudrate.Text !="" && cmb_paritate.Text != "" && cmb_stop_bits.Text !="")
                {
                    ModClient = new ModbusClient(cmb_port.Text)
                    {
                        Baudrate = int.Parse(cmb_baudrate.Text)
                    };
                    if (cmb_stop_bits.Text == "1")
                    {
                        ModClient.StopBits = System.IO.Ports.StopBits.One;
                    }
                    else if (cmb_stop_bits.Text == "2")
                    {
                        ModClient.StopBits = System.IO.Ports.StopBits.Two;
                    }
                    else
                    {
                        ModClient.StopBits = System.IO.Ports.StopBits.None;
                    }

                    if (cmb_paritate.Text == "None")
                    {
                        ModClient.Parity = System.IO.Ports.Parity.None;
                    }
                    else if (cmb_paritate.Text == "Even")
                    {
                        ModClient.Parity = System.IO.Ports.Parity.Even;
                    }
                    else if (cmb_paritate.Text == "Odd")
                    {
                        ModClient.Parity = System.IO.Ports.Parity.Odd;
                    }

                    try
                    {
                        disconnect_connect = true;
                        timer_repeat.Start();
                        ModClient.Connect();
                        connect = true;
                        disconnect = false;
                        btn_deconectare.BackColor = System.Drawing.Color.LightCyan;
                        btn_conectare.BackColor = System.Drawing.Color.Chartreuse;
                    } catch (Exception)
                    {
                        MessageBox.Show("Verificati comunicatia!","Comunicatie",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                else
                {
                    MessageBox.Show("Trebuie sa selectati toate campurile","Date incomplete", MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                MessageBox.Show("Sunteti deja conectat!", "Sunteti deja conectat!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private void Btn_deconectare_Click(object sender, EventArgs e)
        {
            if(disconnect == false) { 
                disconnect_connect = false;
                timer_repeat.Stop();
                ModClient.Disconnect();
                connect = false;
                disconnect = true;
                btn_deconectare.BackColor = System.Drawing.Color.Red;
                btn_conectare.BackColor = System.Drawing.Color.LightCyan;
            }
            else
            {
                MessageBox.Show("Sunteti deja deconectat!", "Sunteti deja deconectat!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        #endregion

        #region Switch button use for activate the Slave
        int btn_switch_orange = 0;
        bool activate_the_slave_I = false;
        private void Activ_dezactiv_slave_orange_Click(object sender, EventArgs e)
        {
            if(txt_id_S_Orange.Text != "" && cmb_fct_S_Orange.Text!="Functie" && txt_first_rgs_S_Orange.Text != ""  && txt_nr_reg_S_orange.Text!="")
            {
                btn_switch_orange++;

                if (btn_switch_orange == 1)
                {
                    activate_the_slave_I = true;
                    for (int i = 0; i < Convert.ToInt32(txt_nr_reg_S_orange.Text); i++)
                    {
                        dta_grd_S_Orange.Rows.Add();
                        dta_grd_S_Orange.Rows[i].Cells[0].Value = Convert.ToInt32(txt_first_rgs_S_Orange.Text) + i;
                    }
                }
                else if (btn_switch_orange == 2)
                {
                    activate_the_slave_I = false;
                    dta_grd_S_Orange.Rows.Clear();
                    btn_switch_orange = 0;
                }
            }
            else
            {
                activ_dezactiv_slave_orange.Switched = true;
            }
        }

        int btn_switch_pink = 0;
        bool activate_the_slave_II = false;
        private void Activ_dezactiv_slave_pink_Click(object sender, EventArgs e)
        {
            if (txt_id_S_Pink.Text != "" && cmb_S_Pink .Text != "Functie" && txt_first_rgs_Pink.Text != "" && txt_nr_reg_S_pink.Text != "")
            {
                btn_switch_pink++;

                if (btn_switch_pink == 1)
                {
                    activate_the_slave_II = true;
                    for (int i = 0; i < Convert.ToInt32(txt_nr_reg_S_pink.Text); i++)
                    {
                        dta_grd_S_pink.Rows.Add();
                        dta_grd_S_pink.Rows[i].Cells[0].Value = Convert.ToInt32(txt_first_rgs_Pink.Text) + i;
                    }
                }
                else if (btn_switch_pink == 2)
                {
                    activate_the_slave_II = false;
                    dta_grd_S_pink.Rows.Clear();
                    btn_switch_pink = 0;
                }
                else
                {
                    btn_switch_pink = 0;
                }
            }
            else
            {
                activ_dezactiv_slave_pink.Switched = true;
            }
        }

        int btn_switch_blue = 0;
        bool activate_the_slave_III = false;
        private void Activ_dezactiv_slave_blue_Click(object sender, EventArgs e)
        {
            if (txt_id_S_Blue.Text != "" && cmb_S_Blue.Text != "Functie" && txt_first_rgs_S_Blue.Text != "" && txt_nr_reg_S_Blue.Text != "") { 
                btn_switch_blue++;

                if (btn_switch_blue == 1)
                {
                    activate_the_slave_III = true;
                    for (int i = 0; i < Convert.ToInt32(txt_nr_reg_S_Blue.Text); i++)
                    {
                        dta_grd_S_blue.Rows.Add();
                        dta_grd_S_blue.Rows[i].Cells[0].Value = Convert.ToInt32(txt_first_rgs_S_Blue.Text) + i;
                    }
                }
                else if (btn_switch_blue == 2)
                {
                    activate_the_slave_III = false;
                    dta_grd_S_blue.Rows.Clear();
                    btn_switch_blue = 0;
                }
               
            }
            else
            {
                activ_dezactiv_slave_blue.Switched = true;
            }
        }
        #endregion

        #region Regex txt Box slave orange
        private void Txt_id_S_Orange_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(txt_id_S_Orange.Text, "[^0-9]"))
            {
                MessageBox.Show("Please enter only numbers.", "Please enter only numbers!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt_id_S_Orange.Text = txt_id_S_Orange.Text.Remove(txt_id_S_Orange.Text.Length - 1);
            }
        }

        private void Txt_first_rgs_S_Orange_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(txt_first_rgs_S_Orange.Text, "[^0-9]"))
            {
                MessageBox.Show("Please enter only numbers.", "Please enter only numbers!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt_first_rgs_S_Orange.Text = txt_first_rgs_S_Orange.Text.Remove(txt_first_rgs_S_Orange.Text.Length - 1);
            }
        }

        private void Txt_nr_reg_S_orange_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(txt_nr_reg_S_orange.Text, "[^0-9]"))
            {
                MessageBox.Show("Please enter only numbers.", "Please enter only numbers!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt_nr_reg_S_orange.Text = txt_nr_reg_S_orange.Text.Remove(txt_nr_reg_S_orange.Text.Length - 1);
            }

            if (txt_nr_reg_S_orange.Text.Length != 0 && Convert.ToInt32(txt_nr_reg_S_orange.Text) > 21)
            {
                MessageBox.Show("Please enter only numbers less than 20.", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt_nr_reg_S_orange.Text = "";
            }

        }
        #endregion

        #region txt box slave pink
        private void Txt_id_S_Pink_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(txt_id_S_Pink.Text, "[^0-9]"))
            {
                MessageBox.Show("Please enter only numbers.", "Please enter only numbers!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt_id_S_Pink.Text = txt_id_S_Pink.Text.Remove(txt_id_S_Pink.Text.Length - 1);
            }
        }

        private void Txt_first_rgs_Pink_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(txt_first_rgs_Pink.Text, "[^0-9]"))
            {
                MessageBox.Show("Please enter only numbers.", "Please enter only numbers!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt_first_rgs_Pink.Text = txt_first_rgs_Pink.Text.Remove(txt_first_rgs_Pink.Text.Length - 1);
            }
        }

        private void Txt_nr_reg_S_pink_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(txt_nr_reg_S_pink.Text, "[^0-9]"))
            {
                MessageBox.Show("Please enter only numbers.", "Please enter only numbers!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt_nr_reg_S_pink.Text = txt_nr_reg_S_pink.Text.Remove(txt_nr_reg_S_pink.Text.Length - 1);

            }

            if (txt_nr_reg_S_pink.Text.Length != 0 && Convert.ToInt32(txt_nr_reg_S_pink.Text) > 21)
            {
                MessageBox.Show("Please enter only numbers less than 20.", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt_nr_reg_S_pink.Text = "";
            }
        }
        #endregion

        #region txt box slave blue
        private void Txt_id_S_Blue_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(txt_id_S_Blue.Text, "[^0-9]"))
            {
                MessageBox.Show("Please enter only numbers.", "Please enter only numbers!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt_id_S_Blue.Text = txt_id_S_Blue.Text.Remove(txt_id_S_Blue.Text.Length - 1);
            }
        }

        private void Txt_first_rgs_S_Blue_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(txt_first_rgs_S_Blue.Text, "[^0-9]"))
            {
                MessageBox.Show("Please enter only numbers.", "Please enter only numbers!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt_first_rgs_S_Blue.Text = txt_first_rgs_S_Blue.Text.Remove(txt_first_rgs_S_Blue.Text.Length - 1);
            }
        }

        private void Txt_nr_reg_S_Blue_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(txt_nr_reg_S_Blue.Text, "[^0-9]"))
            {
                MessageBox.Show("Please enter only numbers.", "Please enter only numbers!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt_nr_reg_S_Blue.Text = txt_nr_reg_S_Blue.Text.Remove(txt_nr_reg_S_Blue.Text.Length - 1);
            }

            if (txt_nr_reg_S_Blue.Text.Length != 0 && Convert.ToInt32(txt_nr_reg_S_Blue.Text) > 21)
            {
                MessageBox.Show("Please enter only numbers less than 20.", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt_nr_reg_S_Blue.Text = "";
            }
        }
        #endregion

        #region Write data
        private void Dta_grd_S_Orange_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (activate_the_slave_I == true)
            {
                if (cmb_fct_S_Orange.Text == "Write Holding")
                {
                    try { 
                        ModClient.WriteSingleRegister(int.Parse(dta_grd_S_Orange.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString()), int.Parse(dta_grd_S_Orange.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()));
                    }
                    catch(Exception ex)
                    {
                        activate_the_slave_I = false;
                        if (ex.Message.ToString() == "Function code not supported by master")
                        {
                            MessageBox.Show("Registrii de tip Holding registers nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                        {
                            MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        btn_switch_orange = 0;
                        dta_grd_S_Orange.Rows.Clear();
                        activ_dezactiv_slave_orange.Switched = false;
                    }
                }

                if (cmb_fct_S_Orange.Text == "Write Coil")
                {
                    string data_string_send = dta_grd_S_Orange.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    string  cell_adr = dta_grd_S_Orange.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString();
                   
                    if (data_string_send == "True" || data_string_send == "true" || data_string_send == "TRUE" || data_string_send == "1")
                    {
                        data_string_send = "true"; 
                        try
                        {
                            ModClient.WriteSingleCoil(int.Parse(cell_adr), bool.Parse(data_string_send));
                        }
                        catch (Exception ex)
                        {
                            activate_the_slave_I = false;
                            if (ex.Message.ToString() == "Function code not supported by master")
                            {
                                MessageBox.Show("Registrii de tip Coils nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                            {
                                MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            btn_switch_orange= 0;
                            dta_grd_S_Orange.Rows.Clear();
                            activ_dezactiv_slave_orange.Switched = false;
                        }
                    }
                    else if(data_string_send == "False" || data_string_send == "false" || data_string_send == "FALSE" || data_string_send =="0")
                    {
                        data_string_send = "false";
                        try
                        {
                            ModClient.WriteSingleCoil(int.Parse(cell_adr), bool.Parse(data_string_send));
                        }
                        catch (Exception ex)
                        {
                            activate_the_slave_I = false;
                            if (ex.Message.ToString() == "Function code not supported by master")
                            {
                                MessageBox.Show("Registrii de tip Coils nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                            {
                                MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            btn_switch_orange = 0;
                            dta_grd_S_Orange.Rows.Clear();
                            activ_dezactiv_slave_orange.Switched = false;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Incercati sa utilizati: True/true/TRUE/1 pentru TRUE si False/false/FALSE/0 pentru FALSE", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
            }
        }

        private void Dta_grd_S_pink_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (activate_the_slave_II == true)
            {
                if (cmb_S_Pink.Text == "Write Holding")
                {
                    try
                    {
                        ModClient.WriteSingleRegister(int.Parse(dta_grd_S_pink.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString()), int.Parse(dta_grd_S_pink.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()));
                    }
                    catch (Exception ex)
                    {
                        activate_the_slave_II = false;
                        if (ex.Message.ToString() == "Function code not supported by master")
                        {
                            MessageBox.Show("Registrii de tip Holding registers nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                        {
                            MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        btn_switch_pink = 0;
                        dta_grd_S_pink.Rows.Clear();
                        activ_dezactiv_slave_pink.Switched = false;
                    }
                }

                if (cmb_S_Pink.Text == "Write Coil")
                {
                    string cell_adr = dta_grd_S_pink.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString();
                    string data_string_send = dta_grd_S_pink.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    //data_string_send = dta_grd_S_pink.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    //cell_adr = dta_grd_S_pink.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString();
                    if (data_string_send == "True" || data_string_send == "true" || data_string_send == "TRUE" || data_string_send == "1")
                    {
                        data_string_send = "true"; 
                        try
                        {
                            ModClient.WriteSingleCoil(int.Parse(cell_adr), bool.Parse(data_string_send));
                        }
                        catch (Exception ex)
                        {
                            activate_the_slave_II = false;
                            if (ex.Message.ToString() == "Function code not supported by master")
                            {
                                MessageBox.Show("Registrii de tip Coils nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                            {
                                MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            btn_switch_pink = 0;
                            dta_grd_S_pink.Rows.Clear();
                            activ_dezactiv_slave_pink.Switched = false;
                        }
                    }
                    else if (data_string_send == "False" || data_string_send == "false" || data_string_send == "FALSE" || data_string_send == "0")
                    {
                        data_string_send = "false";
                        try
                        {
                            ModClient.WriteSingleCoil(int.Parse(cell_adr), bool.Parse(data_string_send));
                        }
                        catch (Exception ex)
                        {
                            activate_the_slave_II = false;
                            if (ex.Message.ToString() == "Function code not supported by master")
                            {
                                MessageBox.Show("Registrii de tip Coils nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                            {
                                MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            btn_switch_pink = 0;
                            dta_grd_S_pink.Rows.Clear();
                            activ_dezactiv_slave_pink.Switched = false;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Incercati sa utilizati: True/true/TRUE/1 pentru TRUE si False/false/FALSE/0 pentru FALSE", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
            }
        }

        private void Dta_grd_S_blue_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (activate_the_slave_III == true)
            {
                if (cmb_S_Blue.Text == "Write Holding")
                {
                    try
                    {
                        ModClient.WriteSingleRegister(int.Parse(dta_grd_S_blue.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString()), int.Parse(dta_grd_S_blue.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()));
                    }
                    catch (Exception ex)
                    {
                        activate_the_slave_III = false;
                        if (ex.Message.ToString() == "Function code not supported by master")
                        {
                            MessageBox.Show("Registrii de tip Holding registers nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                        {
                            MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        btn_switch_blue = 0;
                        dta_grd_S_blue.Rows.Clear();
                        activ_dezactiv_slave_blue.Switched = false;
                    }
                }

                if (cmb_S_Blue.Text == "Write Coil")
                {
                    string cell_adr = dta_grd_S_blue.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString();
                    string data_string_send = dta_grd_S_blue.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    //data_string_send = dta_grd_S_blue.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    //cell_adr = dta_grd_S_blue.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString();
                    if (data_string_send == "True" || data_string_send == "true" || data_string_send == "TRUE" || data_string_send == "1")
                    {
                        data_string_send = "true";
                        try
                        {
                            ModClient.WriteSingleCoil(int.Parse(cell_adr), bool.Parse(data_string_send));
                        }
                        catch (Exception ex)
                        {
                            activate_the_slave_III = false;
                            if (ex.Message.ToString() == "Function code not supported by master")
                            {
                                MessageBox.Show("Registrii de tip Coils nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                            {
                                MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            btn_switch_blue = 0;
                            dta_grd_S_blue.Rows.Clear();
                            activ_dezactiv_slave_blue.Switched = false;
                        }
                    }
                    else if (data_string_send == "False" || data_string_send == "false" || data_string_send == "FALSE" || data_string_send == "0")
                    {
                        data_string_send = "false";
                        try
                        {
                            ModClient.WriteSingleCoil(int.Parse(cell_adr), bool.Parse(data_string_send));
                        }
                        catch (Exception ex)
                        {
                            activate_the_slave_III = false;
                            if (ex.Message.ToString() == "Function code not supported by master")
                            {
                                MessageBox.Show("Registrii de tip Coils nu sunt configurati!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            else if (ex.Message.ToString() == "Starting address invalid or starting address + quantity invalid")
                            {
                                MessageBox.Show("Adresa de start + cantitatea de registrii selectati depaseste numarul de registii existenti in retea!", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            btn_switch_blue = 0;
                            dta_grd_S_blue.Rows.Clear();
                            activ_dezactiv_slave_blue.Switched = false;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Incercati sa utilizati: True/true/TRUE/1 pentru TRUE si False/false/FALSE/0 pentru FALSE", "Atentie!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
            }
        }
        #endregion

    }
}
