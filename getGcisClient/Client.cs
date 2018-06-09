using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Sockets;

namespace getGcisClient
{
    class Client
    {
        public string IPAddr { private set; get; }
        public int port { private set; get; }
        private TcpClient client;

        public Client(string IPAddr,int port)
        {
            this.IPAddr = IPAddr;
            this.port = port;
        }

        public void Connect()
        {
            if (client != null)
                client.Close();
            client = new TcpClient();
            try
            {
                client.Connect(IPAddr, port);
                client.ReceiveBufferSize = 1024;                
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
            CommunicateWithServer();
        }

        private void CommunicateWithServer()
        {
            int reclength = 0;
            NetworkStream netStream = client.GetStream();
            byte[] buffer = new byte[1024];
            while(client.Connected)
            {
                if ((reclength = netStream.Read(buffer,0,buffer.Length)) != 0)
                {
                    string recMessage = Encoding.UTF8.GetString(buffer, 0, reclength);
                    switch(recMessage)
                    {
                        case "added":
                            Console.WriteLine("已連線，等待接受服務...");
                            break;
                        case "ready":
                            Console.WriteLine("開始傳送檔案列表...");
                            break;
                    }
                }
            }
        }
    }
}
