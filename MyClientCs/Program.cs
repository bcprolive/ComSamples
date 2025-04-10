﻿using System;
using System.Threading;
//using MyInterfaces;

namespace MyClientCs
{
    class Program
    {
        static void CommunicateWithServer()
        {
            // create or connect to server object in a separate process
            // equivalent to Activator.CreateInstance(Type.GetTypeFromCLSID(typeof(MyInterfaces.MyServerClass).GUID))
            var server = new MyInterfaces.MyServerClass();

            {
                var cruncher = server.GetNumberCruncher();
                double pi = cruncher.ComputePi();
                Console.WriteLine($"pi = {pi}");

                // release reference to help GC clean up (not strictly needed)
                cruncher = null;
            }
            {
                var callback = new ClientCallback();
                server.Subscribe(callback);

                // wait 5 seconds before exiting to give the server time to send messages
                Thread.Sleep(5000);

                server.Unsubscribe(callback);
                // release reference to help GC clean up (not strictly needed)
                callback = null;
            }

            // release reference to help GC clean up (not strictly needed)
            server = null;
        }

        [MTAThread] // or [STAThread]
        static void Main(string[] _)
        {
            // Perform COM calls in a separate function.
            // Work-around for GC problem mentioned below.
            CommunicateWithServer();

            // Trigger GC to release references seen by server.
            // WARNING: Doesn't clean up properly if called from the same function as COM calls (observed with .Net 7.0)
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }
    }

    /** Non-creatable COM class that doesn't need any CLSID. */
    class ClientCallback : MyInterfaces.IMyClient
    {
        public void XmitMessage(ref MyInterfaces.Message msg)
        {
            Console.WriteLine("Received message:");
            Console.WriteLine("  sev=" + msg.sev);
            Console.WriteLine("  time=" + msg.time);
            Console.WriteLine("  value=" + msg.value);
            Console.WriteLine("  desc=" + msg.desc);
            Console.WriteLine("  color=(" + string.Join(", ", msg.color)+")");
            Console.WriteLine("  data=[" + string.Join(",", msg.data)+"]");
        }
    }
}
