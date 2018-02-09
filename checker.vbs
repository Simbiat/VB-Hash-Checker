'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
' Hash check
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Function hashfilecheck(hashlist)
  if not objFSO.fileexists(hashlist) Then
    call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[HashErr] " & logstartline() & "" & hashlist & " is missing! Administrator needs to run Hash Updater!", hashlist & " is missing!" & vbcrlf & "Administrator needs to run Hash Updater!", ForAppending, 3)
    hashfilecheck=false
    exit function
  end if
  dim flog, failedfiles, itterh, strNextLine, arrServiceList, inFile, hashoutconv, check, fLog2
  failedfiles = ""
  itterh = 1
  Set fLog=objFSO.openTextfile(hashlist)
  Do Until fLog.AtEndOfStream
    strNextLine = fLog.Readline
    if Len(strNextLine)>0 then
      arrServiceList = Split(strNextLine , ",")
      if objFSO.fileexists(arrServiceList(0)) Then
        statusbar.value = "Checking integrity of " & arrServiceList(0)
        set inFile=objFSO.CreateTextfile(hashinfile,true)
        inFile.WriteLine arrServiceList(0)
        inFile.WriteLine arrServiceList(1)
        inFile.Close
        check=oShell.Run(hashfileexe & " -T " & hashinfile & " " & hashoutfile, 0, True)
        if objFSO.fileexists(hashinfile) Then
          objFSO.DeleteFile hashinfile,1
        End if
        Set fLog2=objFSO.openTextfile(hashoutfile, ForReading, false, TristateFalse)
        hashoutconv = ""
        Do Until fLog2.AtEndOfStream
          hashoutconv = hashoutconv & StrConv(fLog2.Readline, "Windows-1251", "cp866") & " ||| "
        Loop
        flog2.close
        if objFSO.fileexists(hashinfile) Then
          objFSO.DeleteFile hashoutfile,1
        End if
        if check=0 then
          call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[HashOK ]" & logstartline() & "" & hashoutconv, arrServiceList(0) & " successfully validated", ForAppending, 0)
        Else
          call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[HashErr]" & logstartline() & "" & hashoutconv, arrServiceList(0) & " failed validation", ForAppending, 0)
          If (itterh Mod 2) = 0 Then
            failedfiles = failedfiles & arrServiceList(0) & vbcrlf
          else
            failedfiles = failedfiles & arrServiceList(0) & "; "
          end if
          itterh = itterh + 1
        End if
      else
        call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[HashWar]" & logstartline() & "" & arrServiceList(0) & " was not found", 1, ForAppending, 0)
      End if
    End If
  Loop
  fLog.close
  if failedfiles = "" then
    hashfilecheck=true
  else
    msgbox "Following files failed validation! Note, that list may be incomplete. For complete list refer to log (HashErr and HashWar message types)" & vbcrlf & vbcrlf & failedfiles, vbCritical, "Error"
    hashfilecheck=false
  end if
End Function
