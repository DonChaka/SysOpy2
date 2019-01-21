set WshNetwork = WScript.CreateObject("WScript.Network")
WshNetwork.MapNetworkDrive "Z:", "\\Server\Share"
WshNetwork.AddWindowsPrinterConnection("\\Server\Printer")

WScript.Echo "Dyski: "
WScript.Echo EnumNetworkDrives

WScript.Echo "Drukarki: "
WScript.Echo WshNetwork.EnumPrinterConnection