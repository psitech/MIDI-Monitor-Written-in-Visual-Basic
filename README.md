# MIDI-Monitor-Written-in-Visual-Basic

This repositiory presents a MIDI monitor written in VB, using Visual Basic Express 2010.

Being curious what MIDI data is exactly outputted by my MIDI keyboard, I decided to write this little tool.

![image](https://user-images.githubusercontent.com/27091013/216694671-dc82a82a-d4ea-416b-97fe-526d875926f8.png)

First the complete code:
``` 
Imports System.Threading
Imports System.Runtime.InteropServices

Public Class Form1

    Public Declare Function midiInGetNumDevs Lib "winmm.dll" () As Integer
    Public Declare Function midiInGetDevCaps Lib "winmm.dll" _
           Alias "midiInGetDevCapsA" (ByVal uDeviceID As Integer, _
           ByRef lpCaps As MIDIINCAPS, ByVal uSize As Integer) As Integer
    Public Declare Function midiInOpen Lib "winmm.dll" _
           (ByRef hMidiIn As Integer, ByVal uDeviceID As Integer, _
           ByVal dwCallback As MidiInCallback, ByVal dwInstance As Integer, _
           ByVal dwFlags As Integer) As Integer
    Public Declare Function midiInStart Lib "winmm.dll" (ByVal hMidiIn As Integer) As Integer
    Public Declare Function midiInStop Lib "winmm.dll" (ByVal hMidiIn As Integer) As Integer
    Public Declare Function midiInReset Lib "winmm.dll" (ByVal hMidiIn As Integer) As Integer
    Public Declare Function midiInClose Lib "winmm.dll" (ByVal hMidiIn As Integer) As Integer

    Public Delegate Function MidiInCallback(ByVal hMidiIn As Integer, _
           ByVal wMsg As UInteger, ByVal dwInstance As Integer, _
           ByVal dwParam1 As Integer, ByVal dwParam2 As Integer) As Integer
    Public ptrCallback As New MidiInCallback(AddressOf MidiInProc)
    Public Const CALLBACK_FUNCTION As Integer = &H30000
    Public Const MIDI_IO_STATUS = &H20

    Public Delegate Sub DisplayDataDelegate(dwParam1)

    Public Structure MIDIINCAPS
        Dim wMid As Int16 ' Manufacturer ID
        Dim wPid As Int16 ' Product ID
        Dim vDriverVersion As Integer ' Driver version
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> _
        Dim szPname As String ' Product Name
        Dim dwSupport As Integer ' Reserved
    End Structure

    Dim hMidiIn As Integer
    Dim StatusByte As Byte
    Dim DataByte1 As Byte
    Dim DataByte2 As Byte
    Dim MonitorActive As Boolean = False
    Dim HideMidiSysMessages As Boolean = False

    Function MidiInProc(ByVal hMidiIn As Integer, _
        ByVal wMsg As UInteger, ByVal dwInstance As Integer, _
        ByVal dwParam1 As Integer, ByVal dwParam2 As Integer) As Integer
        If MonitorActive = True Then
            TextBox1.Invoke(New DisplayDataDelegate(AddressOf DisplayData), _
                            New Object() {dwParam1})
        End If
    End Function

    Private Sub DisplayData(dwParam1)
        If ((HideMidiSysMessages = True) And ((dwParam1 And &HF0) = &HF0)) Then
            Exit Sub
        Else
            StatusByte = (dwParam1 And &HFF)
            DataByte1 = (dwParam1 And &HFF00) >> 8
            DataByte2 = (dwParam1 And &HFF0000) >> 16
            TextBox1.AppendText(String.Format("{0:X2} {1:X2} {2:X2}{3}", _
                                StatusByte, DataByte1, DataByte2, vbCrLf))
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, _
            ByVal e As System.EventArgs) Handles Me.Load
        Me.Show()
        If midiInGetNumDevs() = 0 Then
            MsgBox("No MIDI devices connected")
            Application.Exit()
        End If

        Dim InCaps As New MIDIINCAPS
        Dim DevCnt As Integer

        For DevCnt = 0 To (midiInGetNumDevs - 1)
            midiInGetDevCaps(DevCnt, InCaps, Len(InCaps))
            ComboBox1.Items.Add(InCaps.szPname)
        Next DevCnt
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, _
            e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox1.Enabled = False
        Dim DeviceID As Integer = ComboBox1.SelectedIndex
        midiInOpen(hMidiIn, DeviceID, ptrCallback, 0, CALLBACK_FUNCTION Or MIDI_IO_STATUS)
        midiInStart(hMidiIn)
        MonitorActive = True
        Button2.Text = "Stop monitor"
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) _
            Handles Button1.Click
        TextBox1.Clear()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) _
            Handles Button2.Click
        If MonitorActive = False Then
            midiInStart(hMidiIn)
            MonitorActive = True
            Button2.Text = "Stop monitor"
        Else
            midiInStop(hMidiIn)
            MonitorActive = False
            Button2.Text = "Start monitor"
        End If
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) _
            Handles Button3.Click
        If HideMidiSysMessages = False Then
            HideMidiSysMessages = True
            Button3.Text = "Show System messages"
        Else
            HideMidiSysMessages = False
            Button3.Text = "Hide System messages"
        End If
    End Sub

    Private Sub Form1_FormClosed(ByVal sender As Object, _
            ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        MonitorActive = False
        midiInStop(hMidiIn)
        midiInReset(hMidiIn)
        'midiInClose(hMidiIn)
        Application.Exit()
    End Sub

End Class
```

The source code is rather self-explanatory. You might want to check out the MIDI reference on the MSDN website for in-depth information for details on the different functions.

Below, I highlighted the parts that took me a while to figure out:


```Function MidiInProc(ByVal hMidiIn As Integer, ByVal wMsg As UInteger, _
         ByVal dwInstance As Integer, ByVal dwParam1 As Integer, _
         ByVal dwParam2 As Integer) As Integer
    If MonitorActive = True Then
        TextBox1.Invoke(New DisplayDataDelegate(AddressOf DisplayData), _
                        New Object() {dwParam1})
    End If
End Function
```

This is the callback function which returns the incoming MIDI messages. dwParam1 contains the 4 bytes MIDI data that I am looking for. From MSDN:

![image](https://user-images.githubusercontent.com/27091013/216692410-6fccdd4c-c30e-4cde-8fef-20ce1b71bcf0.png)

From within the callback function, "DisplayDataDelegate" is invoked to display the received MIDI data.
```
Private Sub DisplayData(dwParam1)
    If ((HideMidiSysMessages = True) And ((dwParam1 And &HF0) = &HF0)) Then
        Exit Sub
    Else
        StatusByte = (dwParam1 And &HFF)
        DataByte1 = (dwParam1 And &HFF00) >> 8
        DataByte2 = (dwParam1 And &HFF0000) >> 16
        TextBox1.AppendText(String.Format("{0:X2} {1:X2} {2:X2}{3}", _
                            StatusByte, DataByte1, DataByte2, vbCrLf))
    End If
End Sub
```
This sub formats the dwParam1 bytes and displays them in the textbox.

The HideMidiSysMessages toggle has been added to suppress the continuous "FE" messages (Active Sensing) my MIDI keyboard is generating.

As a side note: I also found that my MIDI keyboard does not generate a Note Off message but a Note On with velocity '00' when releasing a key (see the screendump above).
  
  
```
Public ptrCallback As New MidiInCallback(AddressOf MidiInProc)
```
This line makes a permanent reference to the callback function. If you don't do this, the callback function will be GarbageCollected by .NET.
   
   
```
'midiInClose(hMidiIn)
```
This line is commented out because sometimes midiInClose(hMidiIn) hangs (meaning: it does not return with or without an error, just hangs). The error shown by the debugger is "Argument not specified for parameter 'hMidiIn' of 'Public Shared Function midiInClose(hMidiIn As Integer) As Integer" (?)

The hanging situation only occurs when a lot of MIDI data is generated by the MIDI keyboard while trying to execute midiInClose. I found an interesting and indepth article [*here*](https://groups.google.com/forum/#!topic/mididev/6OUjHutMpEo) on this issue [kudos to "Les"].

My guess is that midiInClose hangs because MidiInProc is still processing incoming MIDI data. I have no idea how to solve or workaround this issue.

But Application.Exit() definitely ends the program without complaints or leaving threads/callback functions running in the background.

If anyone knows a solution for the midiInClose issue, please feel free to comment!

In the ZIP file, you will also find a ready-to-run EXE file (in \bin\Release). It runs on Win XP and Windows 7 32 & 64 bit.
