dotnet build -r win-x86  --no-self-contained

.\bin\dscom32.exe tlbregister <YOUR-REPO-PATH>\comtestdotnet\comtestdotnet.tlb

regsvr32.exe <YOUR-REPO-PATH>\comtestdotnet\bin\Debug\net6.0\win-x86\comtestdotnet.comhost.dll 