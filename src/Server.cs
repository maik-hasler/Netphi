﻿using System.Runtime.InteropServices;

namespace Netphi;

[ComVisible(true)]
[Guid("56986172-C929-44E5-8B67-41B97B7412D8")]
public class Server : IServer
{
    double IServer.ComputePi()
    {
        double sum = 0.0;
        int sign = 1;
        for (int i = 0; i < 1024; ++i)
        {
            sum += sign / (2.0 * i + 1.0);
            sign *= -1;
        }

        return 4.0 * sum;
    }
}
