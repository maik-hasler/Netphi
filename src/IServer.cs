using System.Runtime.InteropServices;

namespace Netphi;

[ComVisible(true)]
[Guid("2D4D3084-E08E-469A-8442-6B3857C84AB9")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IServer
{
    /// <summary>
    /// Compute the value of the constant Pi.
    /// </summary>
    double ComputePi();
}
