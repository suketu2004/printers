<SCRIPT Language="VBScript">

    Sub Window_Onload

        strComputer = "."


        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

        Set colPrinters = objWMIService.ExecQuery("Select * From Win32_Printer")


        For Each objPrinter in colPrinters

            strPrinter = objPrinter.Name

            Set objOption = Document.createElement("OPTION")

            objOption.Text = strprinter

            objOption.Value = strPrinter

            AvailablePrinters.Add(objOption)

        Next

    End Sub


    Sub SetDefault

        strPrinter = AvailablePrinters.Value

        Set WshNetwork = CreateObject("Wscript.Network")

        WshNetwork.SetDefaultPrinter strPrinter

        Msgbox strprinter & " has been set as your default printer."

    End Sub


</SCRIPT>


<select size="5" name="AvailablePrinters"></select><p>

<input type="button" value="Set as Default" onClick="SetDefault">