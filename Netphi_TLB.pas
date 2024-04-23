unit Netphi_TLB;

// ************************************************************************ //
// WARNUNG
// -------
// Die in dieser Datei deklarierten Typen wurden aus Daten einer Typbibliothek
// generiert. Wenn diese Typbibliothek explizit oder indirekt (über eine
// andere Typbibliothek) reimportiert wird oder wenn der Befehl
// 'Aktualisieren' im Typbibliotheks-Editor während des Bearbeitens der
// Typbibliothek aktiviert ist, wird der Inhalt dieser Datei neu generiert und
// alle manuell vorgenommenen Änderungen gehen verloren.                                        
// ************************************************************************ //

// $Rev: 98336 $
// Datei am 23.04.2024 16:14:51 erzeugt aus der unten beschriebenen Typbibliothek.

// ************************************************************************  //
// Typbib.: C:\TARGIS.ERP\Netphi-master\src\Netphi.tlb (1)
// LIBID: {0B71CAEE-1F3E-73FC-7628-E748C6834654}
// LCID: 0
// Hilfedatei: 
// Hilfe-String: 
// Liste der Abhäng.: 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// SYS_KIND: SYS_WIN32
// Fehler:
//   Fehler beim Erzeugen von Palettenbitmap von (TServer) : Server C:\TARGIS.ERP\Netphi-master\src\bin\Debug\net8.0\win-x86\Netphi.comhost.dll enthält keine Symbole
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit muss ohne Typüberprüfung für Zeiger compiliert werden.  
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}

interface

uses Winapi.Windows, System.Classes, System.Variants, System.Win.StdVCL, Vcl.Graphics, Vcl.OleServer, Winapi.ActiveX;
  
// *********************************************************************//
// In der Typbibliothek deklarierte GUIDS. Die folgenden Präfixe werden verwendet:        
//   Typbibliotheken      : LIBID_xxxx                                      
//   CoClasses            : CLASS_xxxx                                      
//   DISPInterfaces       : DIID_xxxx                                       
//   Nicht-DISP-Interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // Haupt- und Nebenversionen der Typbibliothek
  NetphiMajorVersion = 1;
  NetphiMinorVersion = 0;

  LIBID_Netphi: TGUID = '{0B71CAEE-1F3E-73FC-7628-E748C6834654}';

  IID_IServer: TGUID = '{2D4D3084-E08E-469A-8442-6B3857C84AB9}';
  IID__Server: TGUID = '{7CD632CF-A139-46E6-0423-639DB98E0667}';
  CLASS_Server: TGUID = '{56986172-C929-44E5-8B67-41B97B7412D8}';
type

// *********************************************************************//
// Forward-Deklaration von in der Typbibliothek definierten Typen                     
// *********************************************************************//
  IServer = interface;
  _Server = interface;
  _ServerDisp = dispinterface;

// *********************************************************************//
// Deklaration von in der Typbibliothek definierten CoClasses
// (HINWEIS: Hier wird jede CoClass ihrem Standard-Interface zugewiesen)              
// *********************************************************************//
  Server = _Server;


// *********************************************************************//
// Interface: IServer
// Flags:     (256) OleAutomation
// GUID:      {2D4D3084-E08E-469A-8442-6B3857C84AB9}
// *********************************************************************//
  IServer = interface(IUnknown)
    ['{2D4D3084-E08E-469A-8442-6B3857C84AB9}']
    function ComputePi(out pRetVal: Double): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: _Server
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {7CD632CF-A139-46E6-0423-639DB98E0667}
// *********************************************************************//
  _Server = interface(IDispatch)
    ['{7CD632CF-A139-46E6-0423-639DB98E0667}']
  end;

// *********************************************************************//
// DispIntf:  _ServerDisp
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {7CD632CF-A139-46E6-0423-639DB98E0667}
// *********************************************************************//
  _ServerDisp = dispinterface
    ['{7CD632CF-A139-46E6-0423-639DB98E0667}']
  end;

// *********************************************************************//
// Die Klasse CoServer stellt die Methoden Create und CreateRemote zur
// Verfügung, um Instanzen des Standard-Interface _Server, dargestellt
// von CoClass Server, zu erzeugen. Diese Funktionen können
// von einem Client verwendet werden, der die CoClasses automatisieren
// will, die von dieser Typbibliothek dargestellt werden.                                           
// *********************************************************************//
  CoServer = class
    class function Create: _Server;
    class function CreateRemote(const MachineName: string): _Server;
  end;


// *********************************************************************//
// OLE-Server-Proxy-Klassendeklaration
// Server-Objekt     : TServer
// Hilfe-String      : 
// Standard-Interface: _Server
// Def. Intf. DISP?  : No
// Ereignis-Interface: 
// TypeFlags         : (2) CanCreate
// *********************************************************************//
  TServer = class(TOleServer)
  private
    FIntf: _Server;
    function GetDefaultInterface: _Server;
  protected
    procedure InitServerData; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _Server);
    procedure Disconnect; override;
    property DefaultInterface: _Server read GetDefaultInterface;
  published
  end;

procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses System.Win.ComObj;

class function CoServer.Create: _Server;
begin
  Result := CreateComObject(CLASS_Server) as _Server;
end;

class function CoServer.CreateRemote(const MachineName: string): _Server;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Server) as _Server;
end;

procedure TServer.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{56986172-C929-44E5-8B67-41B97B7412D8}';
    IntfIID:   '{7CD632CF-A139-46E6-0423-639DB98E0667}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TServer.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as _Server;
  end;
end;

procedure TServer.ConnectTo(svrIntf: _Server);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TServer.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TServer.GetDefaultInterface: _Server;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface ist NULL. Die Komponente ist nicht mit dem Server verbunden. Sie müssen vor dieser Operation "Connect" oder "ConnectTo" aufrufen');
  Result := FIntf;
end;

constructor TServer.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TServer.Destroy;
begin
  inherited Destroy;
end;

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TServer]);
end;

end.
