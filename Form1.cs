namespace FTC_200_Control
{
    public partial class Form1 : Form
    {
        private Ftc200 ftc = new("COM10");

        public Form1()
        {
            InitializeComponent();
        }

        private string GetData()
        {
            string data = $"Serial: {ftc.GetSerialNumber()}" + Environment.NewLine +
                $"Firmware: {ftc.GetFirmwareVersion()}" + Environment.NewLine +
                $"Tube Preset: {ftc.GetTubePreset()}" + Environment.NewLine +
                $"Minimum HV: {ftc.GetMinimumHv()}" + Environment.NewLine +
                $"Maximum HV: {ftc.GetMaximumHv()}" + Environment.NewLine +
                $"Control voltage for voltage: {ftc.GetControlVoltage()}" + Environment.NewLine +
                $"Voltage set point: {ftc.GetVoltageSetPoint():F1}" + Environment.NewLine +
                $"Minimum emission current: {ftc.GetMinimumEmissionCurrent()}" + Environment.NewLine +
                $"Maximum emission current: {ftc.GetMaximumEmissionCurrent()}" + Environment.NewLine +
                $"Control voltage for emission current: {ftc.GetControlCurrent()}" + Environment.NewLine +
                $"Emission current set point: {ftc.GetEmissionCurrentSetPoint():F1}" + Environment.NewLine +
                $"Monitored voltage: {ftc.GetMonitoredVoltage():F1}" + Environment.NewLine +
                $"Monitored emission current: {ftc.GetMonitoredEmissionCurrent():F1}" + Environment.NewLine +
                $"Interlock status: {ftc.GetInterlockStatus()}" + Environment.NewLine +
                $"High voltage state: {ftc.GetHighVoltageState()}" + Environment.NewLine +
                $"High voltage restore on powerup: {ftc.GetVoltageRestoreOnPowerUp()}" + Environment.NewLine +
                $"PS input voltage: {ftc.GetPSInputVoltage()}";
            return data;
        }

        private void SetButtons(bool state)
        {
            button1.Enabled = state;
            button2.Enabled = state;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SetButtons(false);

            Task.Run(() =>
            {
                string s = GetData();
                Invoke(() => {
                    textBox1.Text = s;
                    SetButtons(true);
                });
            });
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SetButtons(false);

            Task.Run(() =>
            {
                ftc.SetHighVoltageState(!ftc.GetHighVoltageState());
                Thread.Sleep(200);

                string s = GetData();

                Invoke(() =>
                {
                    textBox1.Text = s;
                    SetButtons(true);
                });
            });
        }
    }
}