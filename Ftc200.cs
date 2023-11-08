using System.Drawing;
using System.Globalization;
using System.IO.Ports;
using System.Text;
using static System.Windows.Forms.AxHost;

namespace FTC_200_Control
{
    internal class Ftc200
    {
        private readonly SerialPort _serialPort;
        private readonly CultureInfo _culture = CultureInfo.GetCultureInfo("en_us");

        public Ftc200(string comPort)
        {
            _serialPort = new SerialPort(comPort)
            {
                BaudRate = 9600,
                DataBits = 8,
                StopBits = StopBits.One,
                Parity = Parity.None,
                Handshake = Handshake.None
            };

            // Try to open the serial port
            try
            {
                _serialPort.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Deu ruim filhão.\nTente novamente.");
            }
            finally
            {
                if (_serialPort.IsOpen)
                {
                    _serialPort.RtsEnable = true;
                    _serialPort.DtrEnable = true;

                    // Set write timeout
                    _serialPort.WriteTimeout = 2000;
                    _serialPort.ReadTimeout = 2000;
                    // If we've gotten this far, we are good to communicate.
                }
            }
        }

        #region Métodos para comunicação
        static private byte CalCheckSum(byte[] _PacketData, int PacketLength)
        {
            byte _CheckSumByte = 0x00;

            for (int i = 1; i <= PacketLength; i++)
                _CheckSumByte ^= _PacketData[i];

            return _CheckSumByte;
        }
        private byte[] SerialSendPacket(byte[] packetBytes, int byteCnt, int returnCnt)
        {
            byte[] response = new byte[returnCnt + 3];
            _serialPort.Write(packetBytes, 0, byteCnt);

            Thread.Sleep(50);

            // Ler a resposta
            _serialPort.Read(response, 0, response.Length);

            if (response[3] != 0)
                throw new Exception("Return status != 0");

            if (response[^1] != CalCheckSum(response, response.Length - 2))
                throw new Exception("Checksum inválido");

            if (response[0] == 0x1B && response[1] == packetBytes[1])
                return response;
            else
                throw new Exception("Resposta do dispositivo é inválida");
        }
        #endregion

        public string GetSerialNumber()
        {
            byte[] outgoingPacket = new byte[4];

            // Buscar número de série...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x43; // Get serial number
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1 
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 7);

            // formato do packet de resposta:
            //   0       1          2          3           4...8            9
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2 ... DATA_N] [CHKSUM]

            // Decodifica a resposta em uma string:
            // ignorando os indices 0-3 e 9
            return Encoding.ASCII.GetString(res, 4, 5);
        }

        public string GetFirmwareVersion()
        {
            byte[] outgoingPacket = new byte[4];

            // Buscar número de série...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x44; // Get firmware Version
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1 
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 6);

            // formato do packet de resposta:
            //   0       1          2          3           4...7            8
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2 ... DATA_N] [CHKSUM]

            // Decodifica a resposta em uma string:
            // ignorando os indices 0-3 e 9
            return Encoding.ASCII.GetString(res, 4, 4);
        }

        public byte GetTubePreset()
        {
            byte[] outgoingPacket = new byte[4];
            // Buscar preset do tubo...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x46; // Get tube type preset 
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1 
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 3);

            // formato do packet de resposta:
            //   0       1          2          3        4        5
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [CHKSUM]
            return res[4];
        }

        public byte GetMinimumHv()
        {
            byte[] outgoingPacket = new byte[4];
            // Buscar alta tensão mínima do tubo...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x48; // Get minimum HV (kV)
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1 
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 3);

            // formato do packet de resposta:
            //   0       1          2          3        4       5
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [CHKSUM]
            return res[4];
        }

        public byte GetMaximumHv()
        {
            byte[] outgoingPacket = new byte[4];
            // Buscar alta tensão máxima do tubo...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x4A; // Get maximum voltage
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1 
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 3);

            // formato do packet de resposta:
            //   0       1          2          3        4        5
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [CHKSUM]
            return res[4];
        }

        public byte GetControlVoltage()
        {
            byte[] outgoingPacket = new byte[4];
            // Buscar tensão de controle de voltagem do tubo...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x4C; // Get control voltage for voltage
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1 
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 4);

            // formato do packet de resposta:
            //   0       1          2          3        4        5        6
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [DATA_3] [CHKSUM]
            return res[5];
        }

        public float GetVoltageSetPoint()
        {
            byte[] outgoingPacket = new byte[4];
            // Buscar tensão alvo...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x4E; // Get voltage set point
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 4);

            // formato do packet de resposta:
            //   0       1          2          3        4        5        6
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [DATA_3] [CHKSUM]
            return float.Parse($"{res[4]}.{res[5]}", _culture);
        }

        public ushort GetMinimumEmissionCurrent()
        {
            byte[] outgoingPacket = new byte[4];
            // Buscar corrente de emissão mínima...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x50; // Get minimum emission current
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 4);

            // formato do packet de resposta:
            //   0       1          2          3        4        5        6
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [DATA_3] [CHKSUM]
            return (ushort)((res[4] << 8) + res[5]);
        }

        public ushort GetMaximumEmissionCurrent()
        {
            byte[] outgoingPacket = new byte[4];
            // Buscar corrente de emissão mínima...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x52; // Get maximum emission current
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 4);

            // formato do packet de resposta:
            //   0       1          2          3        4        5        6
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [DATA_3] [CHKSUM]
            return (ushort)((res[4] << 8) + res[5]);
        }

        public byte GetControlCurrent()
        {
            byte[] outgoingPacket = new byte[4];
            // Buscar tensão de controle de corrente do tubo...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x54; // Get control voltage for emission current
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1 
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 4);

            // formato do packet de resposta:
            //   0       1          2          3        4        5        6
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [DATA_3] [CHKSUM]
            return res[5];
        }

        public float GetEmissionCurrentSetPoint()
        {
            byte[] outgoingPacket = new byte[4];
            // Buscar corrente de emissão alvo...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x56; // Get emission current set point
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 5);

            // formato do packet de resposta:
            //   0       1          2          3        4        5        6        7
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [DATA_3] [DATA_4] [CHKSUM]
            ushort wholePortion = (ushort)((res[4] << 8) + res[5]);
            byte decimalPortion = res[6];
            return float.Parse($"{wholePortion}.{decimalPortion}", _culture);
        }

        public float GetMonitoredVoltage()
        {
            byte[] outgoingPacket = new byte[4];
            // Buscar tensão monitorada...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x57; // Get monitored voltage
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 4);

            // formato do packet de resposta:
            //   0       1          2          3        4        5        6
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [DATA_3] [CHKSUM]
            return float.Parse($"{res[4]}.{res[5]}", _culture);
        }

        public float GetMonitoredEmissionCurrent()
        {
            byte[] outgoingPacket = new byte[4];
            // Buscar corrente de emissão monitorada...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x58; // Get emission current set point
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 5);

            // formato do packet de resposta:
            //   0       1          2          3        4        5        6        7
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [DATA_3] [DATA_4] [CHKSUM]
            ushort wholePortion = (ushort)((res[4] << 8) + res[5]);
            byte decimalPortion = res[6];
            return float.Parse($"{wholePortion}.{decimalPortion}", _culture);
        }

        public bool GetInterlockStatus()
        {
            byte[] outgoingPacket = new byte[4];
            // buscar interlock status...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x59; // Get interlock status
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 3);

            // formato do packet de resposta:
            //   0       1          2          3        4        5
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [CHKSUM]
            return res[4] != 0;
        }

        public bool GetHighVoltageState()
        {
            byte[] outgoingPacket = new byte[4];
            // buscar estado da alta tensão...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x5B; // Get high voltage state
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 3);

            // formato do packet de resposta:
            //   0       1          2          3        4        5
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [CHKSUM]
            return res[4] != 0;
        }

        public bool GetVoltageRestoreOnPowerUp()
        {
            byte[] outgoingPacket = new byte[4];
            // buscar restaurar alta tensão no início...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x5D; // Get high voltage restore on power up
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 3);

            // formato do packet de resposta:
            //   0       1          2          3        4        5
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [CHKSUM]
            return res[4] != 0;
        }

        public byte GetPSInputVoltage()
        {
            byte[] outgoingPacket = new byte[4];
            // buscar restaurar alta tensão no início...
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x63; // Get high voltage restore on power up
            outgoingPacket[2] = 1;    // [NUM_BYTES] = 1
            outgoingPacket[3] = CalCheckSum(outgoingPacket, 2);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 3);

            // formato do packet de resposta:
            //   0       1          2          3        4        5
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [CHKSUM]
            return res[4];
        }

        public void SetHighVoltageState(bool state)
        {
            byte[] outgoingPacket = new byte[5];
            // configurar estado da alta tensão...
            // formato do packet:
            //   0       1          2          3        4
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [CHKSUM]
            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x5A;                  // Set high voltage state
            outgoingPacket[2] = 2;                     // [NUM_BYTES] = 2
            outgoingPacket[3] = (byte)(state ? 1 : 0); // [DATA_1] = (0 = off) (1 = on) 
            outgoingPacket[4] = CalCheckSum(outgoingPacket, 3);

            SerialSendPacket(outgoingPacket, outgoingPacket.Length, 3);
        }

        public void SetVoltageSetPoint(double voltageSetPoint)
        {
            double maxSetPoint = GetMaximumHv();
            double minSetPoint = GetMinimumHv();

            if (voltageSetPoint > maxSetPoint || voltageSetPoint < minSetPoint)
            {
                throw new Exception("Tensão fora de parâmetro");
            }

            byte wholeportion = (byte)Math.Truncate(voltageSetPoint);
            byte decimalportion = (byte)(voltageSetPoint % 1);

            byte[] outgoingPacket = new byte[5];

            // configurar set point da alta tensão...
            // formato do packet:
            //   0       1          2          3        4        5
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [CHKSUM]

            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x4D;                  // Set high voltage set point
            outgoingPacket[2] = 3;                     // [NUM_BYTES] = 3
            outgoingPacket[3] = wholeportion;          // [DATA_1] = Whole number portion of set point
            outgoingPacket[4] = decimalportion;        // [DATA_2] = Decimal portion of set point
            outgoingPacket[5] = CalCheckSum(outgoingPacket, 4);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 4);

            // formato do packet de resposta:
            //   0       1          2          3        4         5        6
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2]  [DATA_3] [CHKSUM]

            if (res[4] != wholeportion || res[5] != decimalportion)
            {
                throw new Exception("Tensão retornada é diferente da enviada");
            }
        }

        public void SetEmissionCurrentSetPoint(double voltageSetPoint)
        {
            double maxSetPoint = GetMaximumEmissionCurrent();
            double minSetPoint = GetMinimumEmissionCurrent();

            if (voltageSetPoint > maxSetPoint || voltageSetPoint < minSetPoint)
            {
                throw new Exception("Corrente de emisão fora de parâmetro");
            }

            ushort wholeportion = (ushort)Math.Truncate(voltageSetPoint);
            byte wholeportionhigh = (byte)(wholeportion >> 8);
            byte wholeportionlow = (byte)(wholeportion & 0xff);
            byte decimalportion = (byte)(voltageSetPoint % 1);

            byte[] outgoingPacket = new byte[5];

            // configurar set point da corrente de emissão...
            // formato do packet:
            //   0       1          2          3        4        5
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [CHKSUM]

            outgoingPacket[0] = 0x1B;
            outgoingPacket[1] = 0x55;                  // Set high voltage set point
            outgoingPacket[2] = 3;                     // [NUM_BYTES] = 3
            outgoingPacket[3] = wholeportionhigh;      // [DATA_1] = Whole number portion of set point (high byte)
            outgoingPacket[4] = wholeportionlow;       // [DATA_2] = Whole number portion of set point (low byte)
            outgoingPacket[5] = decimalportion;        // [DATA_3] = Decimal portion of set point
            outgoingPacket[6] = CalCheckSum(outgoingPacket, 4);

            byte[] res = SerialSendPacket(outgoingPacket, outgoingPacket.Length, 5);

            // formato do packet de resposta:
            //   0       1          2          3        4        5        6        7
            // [ESC] [COMMAND] [NUM_BYTES] [DATA_1] [DATA_2] [DATA_3] [DATA_4] [CHKSUM]

            if (res[4] != wholeportionhigh || res[5] != wholeportionlow || res[6] != decimalportion)
            {
                throw new Exception("Corrente retornada é diferente da enviada");
            }
        }
    }
}
