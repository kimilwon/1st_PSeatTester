#define PROGRAM_RUNNING
#define THREAD_RUN

using System;
using System.ComponentModel;
using System.IO.Ports;
using System.Net;
using System.Net.Sockets;
#if THREAD_RUN
using System.Threading;
#endif
using System.Windows.Forms;
//using System.Threading.Tasks;

namespace PSeatTester
{
    public class IOControl
    {
        private SerialPort IOPortToSerial = null;
        private MyInterface mControl = null;
        private IPEndPoint ep;
        private Socket server;
        private EndPoint remoteEP;
        private EndPoint remoteEP2;
        private byte[] rBuffer1 = new byte[1024];
        private __TcpIP__ Board;
        private __TcpIP__ PC;
        private bool SerialInPort = false;
        private ulong[] InData = { 0x0000000000000000, 0x0000000000000000, 0x0000000000000000 };
        private ulong[] InData2 = { 0x0000000000000000, 0x0000000000000000 };
        private byte[,] OutData = { { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 } };
        private float[] CurrData = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

#if !THREAD_RUN
        private Timer timer1 = new Timer();
#endif
        public IOControl()
        {
        }
        public IOControl(MyInterface mControl, __TcpIP__ Board, __TcpIP__ PC)
        {
            this.PC = PC;
            this.Board = Board;

#if !THREAD_RUN
            timer1.Interval = 10;
            timer1.Tick += timer1_tick;
#endif
            this.mControl = mControl;

            //timer1.Enabled = true;
        }

        public void Open()
        {
            if (UdpOpen() == false)
            {
                MessageBox.Show("마이컴 제어용 통신 포트를 오픈하지 못했습니다.");
            }
            else
            {
                UdpCanCommunicationInit();
                UdpRead();
#if THREAD_RUN
                ThreadSetting();
#else
                timer1.Enabled = true;
#endif
            }
        }

        public void Open(string sPort = null, int Speed = 0,bool SerialInPort = false)
        {
            if ((sPort != null) && (sPort != "") && (sPort != string.Empty))
            {
                IOPortToSerial = new SerialPort()
                {
                    BaudRate = 9600,
                    Parity = Parity.None,
                    DataBits = 8,
                    StopBits = StopBits.One,
                    PortName = sPort
                };

                IOPortToSerial.DataReceived += new SerialDataReceivedEventHandler(IOPortReceive);
                IOPortToSerial.Open();
                //timer2.Interval = 10;
                //timer2.Tick += new EventHandler(time2_tick);
                //timer2.Enabled = true;
                this.SerialInPort = SerialInPort;
            }

            if (UdpOpen() == false)
            {
                MessageBox.Show("마이컴 제어용 통신 포트를 오픈하지 못했습니다.");
            }
            else
            {
                UdpCanCommunicationInit();
                UdpRead();
            }
#if !THREAD_RUN
                timer1.Enabled = true;
#else
            ThreadSetting();
#endif
            return;
        }

        private bool UdpOpen()
        {
            bool Flag;
            string strIP;
            int port1;

            Flag = false;
            if ((PC.IP != "") && (PC.IP != null))
            {
                //종점 생성
                strIP = PC.IP;
                IPAddress ip = IPAddress.Parse(strIP);
                port1 = PC.Port;
                ep = new IPEndPoint(ip, port1);

                server = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
                try
                {
                    server.Bind(ep);
                    isOpen = false;
                    isConnection = false;

                    remoteEP = (EndPoint)new IPEndPoint(ip, port1);

                    if ((Board.IP != "") && (PC.IP != null))
                    {
                        remoteEP2 = (EndPoint)new IPEndPoint(IPAddress.Parse(Board.IP), Board.Port);
                        isOpen = true;
                        isConnection = true;
                        Flag = isOpen;
                    }
                }
                catch// (Exception exp)                
                {
                    //MessageBox.Show("I/O Card 와 Bind 가 되지 않습니다. (이더넷 케이블 확인)");
                    isOpen = false;
                    isConnection = false;
                }
                finally
                {
                    Flag = isOpen;
                }
                server.ReceiveBufferSize = 4096;
                server.SendBufferSize = 4096;
            }

            return Flag;
        }

#if THREAD_RUN
        private BackgroundWorker backgroundWorker1 = null;
        private void ThreadSetting()
        {
            backgroundWorker1 = new BackgroundWorker();

            //ReportProgress메소드를 호출하기 위해서 반드시 true로 설정, false일 경우 ReportProgress메소드를 호출하면 exception 발생
            backgroundWorker1.WorkerReportsProgress = true;
            //스레드에서 취소 지원 여부
            backgroundWorker1.WorkerSupportsCancellation = true;
            //스레드가 run시에 호출되는 핸들러 등록
            backgroundWorker1.DoWork += new DoWorkEventHandler(BackgroundWorker1_DoWork);
            // ReportProgress메소드 호출시 호출되는 핸들러 등록
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker1_ProgressChanged);
            // 스레드 완료(종료)시 호출되는 핸들러 동록
            backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker1_RunWorkerCompleted);


            // 스레드가 Busy(즉, run)가 아니라면
            if (backgroundWorker1.IsBusy != true)
            {
                // 스레드 작동!! 아래 함수 호출 시 위에서 bw.DoWork += new DoWorkEventHandler(bw_DoWork); 에 등록한 핸들러가
                // 호출 됩니다.

                backgroundWorker1.RunWorkerAsync();
            }
            return;
        }

        private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //바로 위에서 worker.ReportProgress((i * 10));호출 시 
            // bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged); 등록한 핸들러가 호출 된다고
            // 하였는데요.. 이 부분에서는 기존 Thread에서 처럼 Dispatcher를 이용하지 않아도 됩니다. 
            // 즉 아래처럼!!사용이 가능합니다.
            //this.tbProgress.Text = (e.ProgressPercentage.ToString() + "%");

            // 기존의 Thread클래스에서 아래와 같이 UI 엘리먼트를 갱신하려면
            // Dispatcher.BeginInvoke(delegate() 
            // {
            //        this.tbProgress.Text = (e.ProgressPercentage.ToString() + "%");
            // )};
            //처럼 처리해야 할 것입니다. 그러나 바로 UI 엘리먼트를 업데이트 하고 있죠??
        }


        //스레드의 run함수가 종료될 경우 해당 핸들러가 호출됩니다.
        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            //스레드가 종료한 이유(사용자 취소, 완료, 에러)에 맞쳐 처리하면 됩니다.
            if ((e.Cancelled == true))
            {
            }
            else if (!(e.Error == null))
            {

            }
            else
            {

            }
        }


        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            do
            {
                //CancellationPending 속성이 true로 set되었다면(위에서 CancelAsync 메소드 호출 시 true로 set된다고 하였죠?
                if ((worker.CancellationPending == true))
                {
                    //루프를 break한다.(즉 스레드 run 핸들러를 벗어나겠죠)
                    e.Cancel = true;
                    break;
                }
                else
                {
                    // 이곳에는 스레드에서 처리할 연산을 넣으시면 됩니다.

                    Processing();

                    Thread.Sleep(5);
                    //await Task.Delay(1000);
                    // 스레드 진행상태 보고 - 이 메소드를 호출 시 위에서 
                    // bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged); 등록한 핸들러가 호출 됩니다.
                    worker.ReportProgress(5);
                }
                if (mControl.isExit == true) worker.CancelAsync();
            } while (true);
            //while (ExitFlag == false);
        }
#endif

        public ulong[] GetInData
        {
            get { if (SerialInPort == false) return InData; else return InData2; }      
        }

        public bool isOpen { get; set; }
        //{
        //    get { return isOpen; }
        //}

        public bool isConnection { get; set; }
        //{
        //    get { return isConnection; }
        //}
        public void UdpClose()
        {
            //소켓 닫기
#if PROGRAM_RUNNING
            if (isOpen == true) server.Close();
            isOpen = false;
#endif

            return;
        }


        private void UdpWrite(int addr, byte[] Data, int Length)
        {
            //데이터 입력
#if PROGRAM_RUNNING
            int SendLength;
            //인코딩(byte[])
            byte[] sBuffer = new byte[100];
            //byte[] sBuffer = Encoding.UTF8.GetBytes(data); // data 가 string 이어야 한다.

            //보내기            

            //sBuffer = Encoding.UTF8.GetBytes(data);// data 가 string 이어야 한다.

            /*
            union_r r = new union_r();

            r.Addr = addr;

            sBuffer[0] = r.c1;
            sBuffer[1] = r.c2;
            sBuffer[2] = r.c3;
            sBuffer[3] = r.c4;
            */
            SendLength = 0;
            sBuffer[SendLength++] = (byte)((addr & 0xff000000) >> 24);
            sBuffer[SendLength++] = (byte)((addr & 0x00ff0000) >> 16);
            sBuffer[SendLength++] = (byte)((addr & 0x0000ff00) >> 8);
            sBuffer[SendLength++] = (byte)((addr & 0x000000ff) >> 0);
            sBuffer[SendLength++] = (byte)Length;

            for (int i = 0; i < Length; i++) sBuffer[SendLength++] = Data[i];

            //if ((isOpen[Ch] == true) && (isConnection[Ch] == true))
            if (isOpen == true)
            {
                try
                {
                    server.SendTo(sBuffer, SendLength, SocketFlags.DontRoute, remoteEP2);
                }
                catch
                {
                }
                finally
                {
                }
            }
#endif
            return;
        }
        private void UpdWrite2()
        {
            //byte[] Data1 = { OutData[1, 0], OutData[1, 1], OutData[1, 2], OutData[1, 3], OutData[1, 4], OutData[1, 5], OutData[1, 6], OutData[1, 7] };
            byte[] Data = { OutData[0, 0], OutData[0, 1], OutData[0, 2], OutData[0, 3], OutData[0, 4], OutData[0, 5], OutData[0, 6], OutData[0, 7] };
            //UdpWrite(0x151, Data1, 8);
            UdpWrite(0x101, Data, 8);
            OutPos = 5;
            OutPosFirst = mControl.공용함수.timeGetTimems();
            OutPosLast = mControl.공용함수.timeGetTimems();
            return;
        }

        private void UpdWrite(short Pos)
        {
            byte[] Data = { OutData[Pos + 1, 0], OutData[Pos + 1, 1], OutData[Pos + 1, 2], OutData[Pos + 1, 3], OutData[Pos + 1, 4], OutData[Pos + 1, 5], OutData[Pos + 1, 6], OutData[Pos + 1, 7] };
            int Addr = 0;

            Addr = 0x181 + (Pos * 0x10);
            UdpWrite(Addr, Data, 8);
            OutPos = 5; 
            OutPosFirst = mControl.공용함수.timeGetTimems();
            OutPosLast = mControl.공용함수.timeGetTimems();
            return;
        }

        public void UdpCanCommunicationInit()
        {
#if PROGRAM_RUNNING
            byte[] sBuffer = { 0xfc, 0x03, 0xff };
            if (remoteEP2 != null) server.SendTo(sBuffer, 3, SocketFlags.DontRoute, remoteEP2);
            mControl.공용함수.timedelay(10);
            if (remoteEP2 != null) server.SendTo(sBuffer, 3, SocketFlags.DontRoute, remoteEP2);
            mControl.공용함수.timedelay(10);
            if (remoteEP2 != null) server.SendTo(sBuffer, 3, SocketFlags.DontRoute, remoteEP2);
#endif
            return;
        }

        private string UdpRead()
        {
            //데이터 받기
            string result = "";

            try
            {
#if PROGRAM_RUNNING
                server.BeginReceiveFrom(rBuffer1, 0, rBuffer1.Length, SocketFlags.None, ref remoteEP, new AsyncCallback(ReceiveUdp), remoteEP);
#endif
            }
            catch
            {
            }
            finally
            {
            }
            return result;
        }

        private void ReceiveUdp(IAsyncResult _AR)
        {
            try
            {
                EndPoint remoteEP = new IPEndPoint(IPAddress.Any, 0);
                // 클라이언트로부터 메시지를 받는다.            
                isConnection = true;
                int ReceivedSize = server.EndReceiveFrom(_AR, ref remoteEP);

                if (0 < ReceivedSize)
                {
                    if (mControl.isExit == false)
                    {
                        //CheckUdpData(rBuffer1, 0, ReceivedSize);                                                
                        CheckUdpData(rBuffer1, ReceivedSize);
                    }
                }
                UdpComCheckFirst = mControl.공용함수.timeGetTimems();
                UdpComCheckLast = mControl.공용함수.timeGetTimems();
                //ReceivedSize[0] = 0;
                server.BeginReceiveFrom(rBuffer1, 0, rBuffer1.Length, SocketFlags.None, ref remoteEP, new AsyncCallback(ReceiveUdp), remoteEP);

            }
            catch// (Exception exp)
            {
                //throw exp;
                //MessageBox.Show(exp.Message);                    
            }

            return;
        }

        //private long UdpFirst = 0;
        //private long UdpLast = 0;
        //private bool isOpen = false;
        //private bool isConnection = false;

        private void CheckUdpData(byte[] data, int Length)
        {
            try
            {
                //this.Text = data;

                if (0 < Length)
                {
                    int CanID = 0;
                    int CanID2 = 0;
                    //float Ad;
                    //ushort x = 0;
                    ushort Ad1 = 0;
                    ushort Ad2 = 0;
                    ushort Ad3 = 0;
                    ushort Ad4 = 0;

                    float Data1;
                    float Data2;
                    float Data3;
                    float Data4;

                    int DataLength;
                    ulong[] cData = { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };

                    CanID = (int)(((data[0] & 0xff) << 24) & 0xff000000);
                    CanID |= (int)(((data[1] & 0xff) << 16) & 0x00ff0000);
                    CanID |= (int)(((data[2] & 0xff) << 8) & 0x0000ff00);
                    CanID |= (int)(((data[3] & 0xff) << 0) & 0x000000ff);
                    DataLength = data[4];

                    for (int i = 0; i < DataLength; i++)
                    {
                        if (i < 8)
                            cData[i] = data[i + 5];
                        else break;
                    }
                                       
                    switch (CanID)
                    {
                        case 0x100: // 3232 In                            
                            InData[0] = (cData[4] << 0) | (cData[5] << 8) | (cData[6] << 16) | (cData[7] << 24);
                            //string s = string.Format("{0:X} {1:X} {2:X} {3:X} ", cData[4], cData[5], cData[6], cData[7]);
                            //this.Text = s;
                            break;
                        case 0x182:
                        case 0x183:
                        case 0x192:
                        case 0x193:
                        case 0x1A2:
                        case 0x1A3:
                            //case 0x1B2:
                            //case 0x1B3:
                            Ad1 = (ushort)(((cData[0] & 0x0f) << 8) | (cData[1] << 0));
                            Ad2 = (ushort)(((cData[2] & 0x0f) << 8) | (cData[3] << 0));
                            Ad3 = (ushort)(((cData[4] & 0x0f) << 8) | (cData[5] << 0));
                            Ad4 = (ushort)(((cData[6] & 0x0f) << 8) | (cData[7] << 0));

                            Ad1 = (ushort)Math.Abs(Ad1 - 2130);
                            Ad2 = (ushort)Math.Abs(Ad2 - 2130);
                            Ad3 = (ushort)Math.Abs(Ad3 - 2130);
                            Ad4 = (ushort)Math.Abs(Ad4 - 2130);

                            Data1 = (float)((20F / 2.5) * ((float)Ad1 * (2.5 / 1966)));
                            Data2 = (float)((20F / 2.5) * ((float)Ad2 * (2.5 / 1966)));
                            Data3 = (float)((20F / 2.5) * ((float)Ad3 * (2.5 / 1966)));
                            Data4 = (float)((20F / 2.5) * ((float)Ad4 * (2.5 / 1966)));

                            CanID2 = CanID - 0X182;
                            switch (CanID2)
                            {
                                case 0:
                                    CurrData[0] = (float)Data1;
                                    CurrData[2] = (float)Data2;
                                    CurrData[4] = (float)Data3;
                                    CurrData[6] = (float)Data4;
                                    break;
                                case 1:
                                    break;
                                case 0x10:
                                    CurrData[8] = Math.Abs((float)Data1);
                                    CurrData[10] = Math.Abs((float)Data2);
                                    CurrData[12] = Math.Abs((float)Data3);
                                    CurrData[14] = Math.Abs((float)Data4);
                                    break;
                                case 0x11:
                                    break;
                                case 0x20:
                                    CurrData[16] = Math.Abs((float)Data1);
                                    CurrData[18] = Math.Abs((float)Data2);
                                    //CurrData[20] = Math.Abs((float)Data3 / 10F);
                                    //CurrData[22] = Math.Abs((float)Data4 / 10F);
                                    break;
                                case 0x21:
                                    break;
                                    //case 0x30:
                                    //case 0x31:
                            }

                            //Ad = (float)(x * (5.0 / 4096.0));
                            ////0 V일때 2.6V 가 뜬다고 함

                            //if (2.6 <= CurrData)
                            //{
                            //    Ad = (float)((CurrData - 2.6) * (30.0 / 2.4)); //2.5V 30A
                            //}
                            //else
                            //{
                            //    Ad = (float)((2.6 - CurrData) * (-30.0 / 2.4)); //2.5V 30A
                            //}
                            //CurrData = Ad;
                            break;
                    }

                    //Application.DoEvents() 를 사용하면 에러 발생 (이 함수고 Callback 으로 호출되어서 그런다.
                    //Application.DoEvents();                    
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message + "\n" + exp.StackTrace);
            }

            return;
        }
        public static int OutPos = 0;
        private long OutPosFirst;
        private long OutPosLast;

        public float[] ADRead
        {
            get { return CurrData; }
        }

        public void outportb(int Out, bool OnOff)
        {
            byte Data = 0x00;

            int Pos = Out / 8;
            int dPos = Out % 8;

            Data = (byte)(0x01 << dPos);

            if (OnOff == true)
                OutData[0, Pos] |= Data;
            else OutData[0, Pos] &= (byte)(~Data);
            //OutPos = 1;
            UpdWrite2();
            return;
        }

        public bool GetOutputCheck(int Out)
        {
            byte Data = 0x00;

            int Pos = Out / 8;
            int dPos = Out % 8;

            Data = (byte)(0x01 << dPos);
            
            if ((OutData[0, Pos] & Data) == Data)
                return true;
            else return false;
        }

        public void Function_outportb(short Out, bool OnOff)
        {
            byte Data = 0x00;

            short IOOut = (short)(Out % 16); //보드 하나당 8 접점 16Bit 지원 (제품 6개 핀 번호 지원)
            short Card = (short)(Out / 16);

            int Pos = (int)IOOut / 8; //보드 하나당 8 접점 16Bit 지원 (제품 6개 핀 번호 지원)
            int dPos = (int)IOOut % 8;

            Data = (byte)(0x01 << dPos);

            if (OnOff == true)
                OutData[Card + 1, Pos] |= Data;
            else OutData[Card + 1, Pos] &= (byte)(~Data);
            //OutPos = 1;
            UpdWrite(Card);
            return;
        }

        public void IOInit()
        {
            for (int i = 0; i < 8; i++)
            {
                OutData[0, i] = 0x00;
                OutData[1, i] = 0x00;
                OutData[2, i] = 0x00;
                OutData[3, i] = 0x00;
            }
            return;
        }

        public void FunctionIOInit()
        {
            for (int i = 0; i < 8; i++)
            {
                OutData[1, i] = 0x00;
                OutData[2, i] = 0x00;
                OutData[3, i] = 0x00;
            }
            return;
        }


        /// <summary>
        /// P32C32 지정된 포트의 I/O 위치를 알아낸다.
        /// </summary>
        /// <param name="Pos"></param>
        /// <returns></returns>
        public __IOData__ IOCheck(short Pos)
        {
            __IOData__ value = new __IOData__();

            int OPos = (int)Pos / 8;
            byte Data = (byte)(0x01 << ((int)Pos % 8));

            value.Card = (short)(OPos / 8);
            value.Pos = (short)(OPos % 8);
            value.Data = Data;

            return value;
        }

        private long UdpComCheckFirst;
        private long UdpComCheckLast;

#if THREAD_RUN
        private void Processing()
        {
            try
            {
                if (isOpen == true)
                {
                    UdpComCheckLast = mControl.공용함수.timeGetTimems();
                    if (isConnection == true)
                    {
                        //컨넥션이 되었을때 통신이 1.5초 이상 이루어 지지 않으면 컨넥션이 끊어진 것으로 처리한다.
                        if (1500 <= (UdpComCheckLast - UdpComCheckFirst))
                        {
                            isConnection = false;
                            UdpComCheckFirst = mControl.공용함수.timeGetTimems();
                            UdpComCheckLast = mControl.공용함수.timeGetTimems();
                        }
                    }
                    else
                    {
                        if (500 <= (UdpComCheckLast - UdpComCheckFirst))
                        {
                            UdpCanCommunicationInit();
                            UdpComCheckFirst = mControl.공용함수.timeGetTimems();
                            UdpComCheckLast = mControl.공용함수.timeGetTimems();
                        }
                    }
                }
                if (OutPos == 1)
                {
                    //byte[] Data1 = { OutData[1, 0], OutData[1, 1], OutData[1, 2], OutData[1, 3], OutData[1, 4], OutData[1, 5], OutData[1, 6], OutData[1, 7] };
                    byte[] Data = { OutData[0, 0], OutData[0, 1], OutData[0, 2], OutData[0, 3], OutData[0, 4], OutData[0, 5], OutData[0, 6], OutData[0, 7] };
                    //UdpWrite(0x151, Data1, 8);
                    UdpWrite(0x101, Data, 8);
                    OutPos = 2;
                    OutPosFirst = mControl.공용함수.timeGetTimems();
                    OutPosLast = mControl.공용함수.timeGetTimems();
                }
                else if (OutPos == 2)
                {
                    byte[] Data = { OutData[1, 0], OutData[1, 1], OutData[1, 2], OutData[1, 3], OutData[1, 4], OutData[1, 5], OutData[1, 6], OutData[1, 7] };
                    UdpWrite(0x181, Data, 8);
                    OutPos = 3;
                    OutPosFirst = mControl.공용함수.timeGetTimems();
                    OutPosLast = mControl.공용함수.timeGetTimems();
                }
                else if (OutPos == 3)
                {
                    byte[] Data = { OutData[2, 0], OutData[2, 1], OutData[2, 2], OutData[2, 3], OutData[2, 4], OutData[2, 5], OutData[2, 6], OutData[2, 7] };
                    UdpWrite(0x191, Data, 8);
                    OutPos = 4;
                    OutPosFirst = mControl.공용함수.timeGetTimems();
                    OutPosLast = mControl.공용함수.timeGetTimems();
                }
                else if (OutPos == 4)
                {
                    byte[] Data = { OutData[3, 0], OutData[3, 1], OutData[3, 2], OutData[3, 3], OutData[3, 4], OutData[3, 5], OutData[3, 6], OutData[3, 7] };
                    UdpWrite(0x1a1, Data, 8);
                    OutPos = 5;
                    OutPosFirst = mControl.공용함수.timeGetTimems();
                    OutPosLast = mControl.공용함수.timeGetTimems();
                }
                else if (OutPos == 5)
                {
                    OutPosFirst = mControl.공용함수.timeGetTimems();
                    OutPosLast = mControl.공용함수.timeGetTimems();
                    OutPos = 0;
                }
                else if (OutPos == 0)
                {
                    OutPosLast = mControl.공용함수.timeGetTimems();
                    if (100 <= (OutPosLast - OutPosFirst))
                    {
                        OutPos = 1;
                    }
                }
                else
                {
                    byte[] Data = { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
                    if (DarkCurrStart == true) Data[0] |= 0x01;
                    UdpWrite(0x140, Data, 8);
                    OutPos = 0;
                    OutPosFirst = mControl.공용함수.timeGetTimems();
                    OutPosLast = mControl.공용함수.timeGetTimems();
                }


                if (IOPortToSerial != null)
                {
                    if (IOPortToSerial.IsOpen == true)
                    {
                        if (SerialSendFlag == false)
                        {
                            SerialPortOut();
                            SerialSendFlag = true;
                            SerialSendFirst = mControl.공용함수.timeGetTimems();
                            SerialSendLast = mControl.공용함수.timeGetTimems();
                        }
                        else
                        {
                            SerialSendLast = mControl.공용함수.timeGetTimems();
                            if (200 <= (SerialSendLast - SerialSendFirst)) SerialSendFlag = false;
                        }
                    }
                }
            }
            catch 
            { 
            }
            finally
            {
            }
        }

        private long SerialSendFirst = 0;
        private long SerialSendLast = 0;
#else
        private void timer1_tick(object sender, EventArgs e)
        {
            try
            {
                timer1.Enabled = false;

                if (isOpen == true)
                {
                    UdpComCheckLast = mControl.공용함수.timeGetTimems();
                    if (isConnection == true)
                    {
                        if (1500 <= (UdpComCheckLast - UdpComCheckFirst))
                        {
                            isConnection = false;
                            UdpComCheckFirst = mControl.공용함수.timeGetTimems();
                            UdpComCheckLast = mControl.공용함수.timeGetTimems();
                        }
                    }
                    else
                    {
                        if (500 <= (UdpComCheckLast - UdpComCheckFirst))
                        {
                            UdpCanCommunicationInit();
                            UdpComCheckFirst = mControl.공용함수.timeGetTimems();
                            UdpComCheckLast = mControl.공용함수.timeGetTimems();
                        }
                    }
                }
                if (OutPos == 1)
                {
                    //byte[] Data1 = { OutData[1, 0], OutData[1, 1], OutData[1, 2], OutData[1, 3], OutData[1, 4], OutData[1, 5], OutData[1, 6], OutData[1, 7] };
                    byte[] Data = { OutData[0, 0], OutData[0, 1], OutData[0, 2], OutData[0, 3], OutData[0, 4], OutData[0, 5], OutData[0, 6], OutData[0, 7] };
                    //UdpWrite(0x151, Data1, 8);
                    UdpWrite(0x101, Data, 8);
                    OutPos = 2;
                    OutPosFirst = mControl.공용함수.timeGetTimems();
                    OutPosLast = mControl.공용함수.timeGetTimems();
                }
                else if (OutPos == 2)
                {
                    byte[] Data = { OutData[1, 0], OutData[1, 1], OutData[1, 2], OutData[1, 3], OutData[1, 4], OutData[1, 5], OutData[1, 6], OutData[1, 7] };
                    UdpWrite(0x181, Data, 8);
                    OutPos = 3;
                    OutPosFirst = mControl.공용함수.timeGetTimems();
                    OutPosLast = mControl.공용함수.timeGetTimems();
                }
                else if (OutPos == 3)
                {
                    byte[] Data = { OutData[2, 0], OutData[2, 1], OutData[2, 2], OutData[2, 3], OutData[2, 4], OutData[2, 5], OutData[2, 6], OutData[2, 7] };
                    UdpWrite(0x191, Data, 8);
                    OutPos = 4;
                    OutPosFirst = mControl.공용함수.timeGetTimems();
                    OutPosLast = mControl.공용함수.timeGetTimems();
                }
                else if (OutPos == 4)
                {
                    byte[] Data = { OutData[3, 0], OutData[3, 1], OutData[3, 2], OutData[3, 3], OutData[3, 4], OutData[3, 5], OutData[3, 6], OutData[3, 7] };
                    UdpWrite(0x1a1, Data, 8);
                    OutPos = 5;
                    OutPosFirst = mControl.공용함수.timeGetTimems();
                    OutPosLast = mControl.공용함수.timeGetTimems();
                }
                else if (OutPos == 5)
                {
                    OutPosFirst = mControl.공용함수.timeGetTimems();
                    OutPosLast = mControl.공용함수.timeGetTimems();
                    OutPos = 0;
                }
                else if (OutPos == 0)
                {
                    OutPosLast = mControl.공용함수.timeGetTimems();
                    if (100 <= (OutPosLast - OutPosFirst))
                    {
                        OutPos = 1;
                    }
                }
                else
                {
                    byte[] Data = { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
                    if (DarkCurrStart == true) Data[0] |= 0x01;
                    UdpWrite(0x140, Data, 8);
                    OutPos = 0;
                    OutPosFirst = mControl.공용함수.timeGetTimems();
                    OutPosLast = mControl.공용함수.timeGetTimems();
                }
            }
            catch { }
            finally
            {
                timer1.Enabled = !mControl.isExit;
            }
        }
#endif
        private bool DarkCurrStart = false;
        public bool DarkCurrentReadStart
        {
            get
            {
                return DarkCurrStart;
            }
            set
            {
                if (DarkCurrStart != value)
                {
                    DarkCurrStart = value;
                    OutPos = 3;
                }
            }
        }

        public void PinSelectToBattOnOff(short Port, bool OnOff)
        {
#if PROGRAM_RUNNING

            try
            {
                try
                {
                    short Pos;

                    Pos = (short)((Port * 2) + 0);
                    Function_outportb(Pos, OnOff);
                }
                catch (Exception Msg)
                {
                    MessageBox.Show(Msg.Message + "\n" + Msg.StackTrace);
                }
            }
            finally
            {
            }
#endif
            return;
        }
        public void PinSelectToGndOnOff(short Port, bool OnOff)
        {
#if PROGRAM_RUNNING

            try
            {
                try
                {
                    short Pos;

                    Pos = (short)((Port * 2) + 1);
                    Function_outportb(Pos, OnOff);
                }
                catch (Exception Msg)
                {
                    MessageBox.Show(Msg.Message + "\n" + Msg.StackTrace);
                }
            }
            finally
            {
            }
#endif
            return;
        }

        public bool GetStartSw
        {
            get
            {
                //__IOData__ Pos = IOCheck(IO_IN.PASS);
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.PASS;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.PASS;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }
        public bool GetResetSw
        {
            get
            {
                //__IOData__ Pos = IOCheck(IO_IN.RESET);

                //ulong Data = (ulong)Pos.Data << Pos.Pos;
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.RESET;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.RESET;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }

            }
        }
        public bool GetAuto
        {
            get
            {
                //__IOData__ Pos = IOCheck(IO_IN.AUTO);

                //ulong Data = (ulong)Pos.Data << Pos.Pos;
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.AUTO;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.AUTO;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }

        public bool GetJigUp
        {
            get
            {
                //__IOData__ Pos = IOCheck(IO_IN.JIG_UP);

                //ulong Data = (ulong)Pos.Data << Pos.Pos;
                ulong Data = (ulong)0x01 << IO_IN.JIG_UP;

                if ((InData[0] & Data) == Data)
                    return true;
                else return false;
            }
        }

        public bool GetProductIn
        {
            get
            {
                //__IOData__ Pos = IOCheck(IO_IN.PRODUCT);

                //ulong Data = (ulong)Pos.Data << Pos.Pos;

                ulong Data = (ulong)0x01 << IO_IN.PRODUCT;

                if ((InData[0] & Data) == Data)
                    return true;
                else return false;
            }
        }

        public bool GetRHSelect
        {
            get
            {
                //__IOData__ Pos = IOCheck(IO_IN.LH_SELECT);

                //ulong Data = (ulong)Pos.Data << Pos.Pos;
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.RH_SELECT;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.RH_SELECT;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }

        public bool GetRHDSelect
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.LHD_RHD;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.RHD;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }
        public bool GetLHDSelect
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.LHD_RHD;

                    if ((InData[0] & Data) != Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.LHD;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }

        public bool GetLumber2Way4WaySelect
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.LUMBER_2WAY;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.LUMBER_2WAY;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }

       
        public bool GetIMSSet
        {
            get
            {
                //__IOData__ Pos = IOCheck(IO_IN.IMS_SET_SW);

                //ulong Data = (ulong)Pos.Data << Pos.Pos;
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.IMS_SET_SW;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.IMS_SET_SW;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }

        public bool GetM1Sw
        {
            get
            {
                //__IOData__ Pos = IOCheck(IO_IN.IMS_M1_SW);

                //ulong Data = (ulong)Pos.Data << Pos.Pos;
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.IMS_M1_SW;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.IMS_M1_SW;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }
        public bool GetM2Sw
        {
            get
            {
                //__IOData__ Pos = IOCheck(IO_IN.IMS_M2_SW);

                //ulong Data = (ulong)Pos.Data << Pos.Pos;
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.IMS_M2_SW;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.IMS_M2_SW;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }

        public bool GetSeatPower
        {
            get 
            {
                if (SerialInPort == false)
                {
                    ulong Data1 = (ulong)0x01 << IO_IN.SEAT_POWER;

                    if ((InData[0] & Data1) == Data1)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data1 = (ulong)0x01 << IO_IN2.OPTION_POWER;

                    if ((InData2[0] & Data1) == Data1)
                        return true;
                    else return false;
                }
            }
        }
        public bool GetWalkIn
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.OPTION_WALKIN;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.OPTION_WALKIN;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }

        public bool GetIMS
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.OPTION_IMS;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.OPTION_IMS;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }

       

        public bool GetSlideFwd
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.SLIDE_FWD;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.SLIDE_FWD;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }
        public bool GetSlideBwd
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.SLIDE_BWD;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.SLIDE_BWD;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }
        public bool GetReclinerFwd
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.RECLINE_FWD;
                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.RECLINE_FWD;
                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }
        public bool GetReclinerBwd
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.RECLINE_BWD;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.RECLINE_BWD;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }
        public bool GetTiltUp
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.TILT_UP;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.TILT_UP;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }
        public bool GetTiltDn
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.TILT_DN;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.TILT_DN;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }

        public bool GetHeightUp
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.HEIGHT_UP;
                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.HEIGHT_UP;
                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }

        public bool GetHeightDn
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.HEIGHT_DN;
                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.HEIGHT_DN;
                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }

        public bool GetLumberFwd
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.LUMBER_FWD;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.LUMBER_FWD;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }
        public bool GetLumberBwd
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.LUMBER_BWD;
                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.LUMBER_BWD;
                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }

        public bool GetLumberUp
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.LUMBER_UP;

                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.LUMBER_UP;

                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }
        public bool GetLumberDn
        {
            get
            {
                if (SerialInPort == false)
                {
                    ulong Data = (ulong)0x01 << IO_IN.LUMBER_DN;
                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.LUMBER_DN;
                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }

        public bool YellowLampOnOff
        {
            set
            {
                outportb(IO_OUT.YELLOW_LAMP, value);
                UpdWrite2();                   
            }
            get
            {
                return GetOutputCheck(IO_OUT.YELLOW_LAMP);
            }
        }
        public bool RedLampOnOff
        {
            set
            {
                outportb(IO_OUT.RED_LAMP, value);
                UpdWrite2();
            }
            get
            {
                return GetOutputCheck(IO_OUT.RED_LAMP);
            }
        }

        public bool GreenLampOnOff
        {
            set
            {
                outportb(IO_OUT.GREEN_LAMP, value);
                UpdWrite2();
            }
            get
            {
                return GetOutputCheck(IO_OUT.GREEN_LAMP);
            }
        }

        public bool ProductInOut
        {
            set
            {
                outportb(IO_OUT.PRODUCT_IN, value);
                UpdWrite2();
            }
            get
            {
                return GetOutputCheck(IO_OUT.PRODUCT_IN);
            }
        }

        public bool SetPinConnection
        {
            set
            {
                outportb(IO_OUT.PIN_CONNECTION, value);
                UpdWrite2();
            }
            get
            {
                return GetOutputCheck(IO_OUT.PIN_CONNECTION);
            }
        }

        public bool TestOKOnOff
        {
            set
            {
                outportb(IO_OUT.TEST_OK, value);
                UpdWrite2();
            }
            get
            {
                return GetOutputCheck(IO_OUT.TEST_OK);
            }
        }
        //public bool SetJigDown
        //{
        //    set
        //    {
        //        outportb(IO_OUT.TEST_JIG_DOWN, value);
        //        UpdWrite2();
        //    }
        //    get
        //    {
        //        return GetOutputCheck(IO_OUT.TEST_JIG_DOWN);
        //    }
        //}
        public bool SetDoorOpen
        {
            set
            {
                outportb(IO_OUT.DOOR_OPEN, value);
                UpdWrite2();
            }
            get
            {
                return GetOutputCheck(IO_OUT.DOOR_OPEN);
            }
        }

        public bool TestINGOnOff
        {
            set
            {
                outportb(IO_OUT.TEST_ING, value);
                UpdWrite2();
            }
            get
            {
                return GetOutputCheck(IO_OUT.TEST_ING);
            }
        }
        //public bool JigOrgOnOff
        //{
        //    set
        //    {
        //        outportb(IO_OUT.JIG_ORG, value);
        //        UpdWrite2();
        //    }
        //    get
        //    {
        //        return GetOutputCheck(IO_OUT.JIG_ORG);
        //    }
        //}

        public bool BuzzerOnOff
        {
            set
            {
                outportb(IO_OUT.BUZZER, value);
                UpdWrite2();
            }
            get
            {
                return GetOutputCheck(IO_OUT.BUZZER);
            }
        }

        public bool GetSlideMidSensor
        {
            get
            {
                ulong Data = (ulong)0x01 << IO_IN.SlideDeliveryPosSensor;

                if ((InData[0] & Data) == Data)
                    return true;
                else return false;
            }           
        }

        public bool GetReclineMidSensor
        {
            get
            {                
                ulong Data = (ulong)0x01 << IO_IN.ReclineDeliveryPosSensor;

                if ((InData[0] & Data) != Data)
                    return true;
                else return false;
            }
        }

        public bool GetPinConnectionSw
        {
            get
            {
                if (SerialInPort == true)
                {
                    ulong Data = (ulong)0x01 << IO_IN.PIN_CONNECTION_SW;
                    if ((InData[0] & Data) == Data)
                        return true;
                    else return false;
                }
                else
                {
                    ulong Data = (ulong)0x01 << IO_IN2.PIN_CONNECTION_SW;
                    if ((InData2[0] & Data) == Data)
                        return true;
                    else return false;
                }
            }
        }
        public bool GetPinConnectionFwd
        {
            get
            {
                ulong Data = (ulong)0x01 << IO_IN.PIN_CONNECTION_FWD;
                if ((InData[0] & Data) == Data)
                    return true;
                else return false;
            }
        }
        public bool GetPinConnectionBwd
        {
            get
            {
                ulong Data = (ulong)0x01 << IO_IN.PIN_CONNECTION_BWD;
                if ((InData[0] & Data) == Data)
                    return true;
                else return false;
            }
        }

        public byte[,] GetOutData
        {
            get { return OutData; }
        }
        private ushort OutData2 = 0x0000;
        private ushort OutDataOld2 = 0x0000;
        private void SerialPortOut()
        {
            Int64 temp_data;

            int temp_data3;
            int crc16;

            int temp_byte;

            byte[] temp_data2 = new byte[8];
            byte[] temp_data4 = new byte[35];

            temp_data = (long)OutData2;

            temp_data3 = (int)(temp_data) & 0xff;
            temp_data2[0] = (byte)temp_data3;

            temp_data3 = (int)(temp_data >> 8) & 0xff;
            temp_data2[1] = (byte)temp_data3;

            temp_data3 = (int)(temp_data >> 16) & 0xff;
            temp_data2[2] = (byte)temp_data3;

            temp_data3 = (int)(temp_data >> 32) & 0xff;
            temp_data2[3] = (byte)temp_data3;


            temp_byte = (byte)temp_data2[0] >> 4;
            temp_data4[0] = hex2ascii((byte)temp_byte);

            temp_byte = (byte)temp_data2[0] & 0xf;
            temp_data4[1] = hex2ascii((byte)temp_byte);

            temp_byte = (byte)temp_data2[1] >> 4;
            temp_data4[2] = hex2ascii((byte)temp_byte);

            temp_byte = (byte)temp_data2[1] & 0xf;
            temp_data4[3] = hex2ascii((byte)temp_byte);

            temp_byte = (byte)temp_data2[2] >> 4;
            temp_data4[4] = hex2ascii((byte)temp_byte);

            temp_byte = (byte)temp_data2[2] & 0xf;
            temp_data4[5] = hex2ascii((byte)temp_byte);

            temp_byte = (byte)temp_data2[3] >> 4;
            temp_data4[6] = hex2ascii((byte)temp_byte);

            temp_byte = (byte)temp_data2[3] & 0xf;
            temp_data4[7] = hex2ascii((byte)temp_byte);

            crc16 = CRC16(temp_data4, 7);

            temp_byte = (byte)(crc16 >> 12) & 0xf;
            temp_data4[8] = hex2ascii((byte)temp_byte);

            temp_byte = (byte)(crc16 >> 8) & 0xf;
            temp_data4[9] = hex2ascii((byte)temp_byte);

            temp_byte = (byte)(crc16 >> 4) & 0xf;
            temp_data4[10] = hex2ascii((byte)temp_byte);

            temp_byte = (byte)crc16 & 0xf;
            temp_data4[11] = hex2ascii((byte)temp_byte);

            if (IOPortToSerial.IsOpen == true) IOPortToSerial.Write(temp_data4, 0, 12);

            OutDataOld2 = OutData2;
            return;
        }


        private void IOPortReceive(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                mControl.공용함수.timedelay(20);
                int Length = IOPortToSerial.BytesToRead;
                byte[] buffer = new byte[Length + 10];

                IOPortToSerial.Read(buffer, 0, Length);

                int temp;
                int crc16;
                int crc16_c;
                byte[] buff2_buf = new byte[53];

                if (Length == 16)
                {

                    temp = Ascii2Hex(buffer[0]) << 4 | Ascii2Hex(buffer[1]);
                    buff2_buf[0] = (byte)temp;

                    temp = Ascii2Hex(buffer[2]) << 4 | Ascii2Hex(buffer[3]);
                    buff2_buf[1] = (byte)temp;

                    temp = Ascii2Hex(buffer[4]) << 4 | Ascii2Hex(buffer[5]);
                    buff2_buf[2] = (byte)temp;

                    temp = Ascii2Hex(buffer[6]) << 4 | Ascii2Hex(buffer[7]);
                    buff2_buf[3] = (byte)temp;

                    temp = Ascii2Hex(buffer[8]) << 4 | Ascii2Hex(buffer[9]);
                    buff2_buf[4] = (byte)temp;

                    temp = Ascii2Hex(buffer[10]) << 4 | Ascii2Hex(buffer[11]);
                    buff2_buf[5] = (byte)temp;

                    crc16 = Ascii2Hex(buffer[12]) << 12 | Ascii2Hex(buffer[13]) << 8 | Ascii2Hex(buffer[14]) << 4 | Ascii2Hex(buffer[15]);


                    crc16_c = CRC16(buffer, 12);
                    if (crc16 == crc16_c)
                    {
                        InData2[0] = ((ulong)buff2_buf[0] << 0) | ((ulong)buff2_buf[1] << 8) | ((ulong)buff2_buf[2] << 16) | ((ulong)buff2_buf[3] << 24);
                        InData2[1] = ((ulong)buff2_buf[4] << 0) | ((ulong)buff2_buf[5] << 8);
                    }

                }
                IOPortToSerial.DiscardInBuffer();
                SerialSendFlag = false;
            }
            catch
            {

            }
            finally
            {
            }
            return;
        }
        private bool SerialSendFlag = false;
        private const UInt16 POLYNORMIAL = 0xA001;

        private ushort CRC16(byte[] bytes, int usDataLen)
        {
            ushort crc = 0xffff, flag, ct = 0;

            while (usDataLen != 0)
            {
                crc ^= bytes[ct];
                for (int i = 0; i < 8; i++)
                {
                    flag = 0;
                    if ((crc & 1) == 1) flag = 1;
                    crc >>= 1;
                    if (flag == 1) crc ^= POLYNORMIAL;
                }
                ct++;
                usDataLen--;
            }
            return crc;
        }

        private byte hex2ascii(byte toconv)
        {
            if (toconv < 0x0A) toconv += 0x30;
            else toconv += 0x37;
            return (toconv);
        }

        private int Ascii2Hex(byte asc)
        {
            int hex;

            if (asc >= '0' && asc <= '9')
            {
                hex = asc - 0x30;
            }
            else if (asc >= 'A' && asc <= 'Z')
            {
                hex = asc - 0x37;
            }
            else
            {
                hex = asc - 0x57;
            }
            return hex;
        }

        ~IOControl()
        {

        }
    }
}
