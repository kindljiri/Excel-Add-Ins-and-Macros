'Decimal To Binary
' =================
' Source: http://groups.google.ca/group/comp.lang.visual.basic/browse_thread/thread/28affecddaca98b4/979c5e918fad7e63
' Author: Randy Birch (MVP Visual Basic)
' NOTE: You can limit the size of the returned
'              answer by specifying the number of bits
Private Function Dec2Bin(ByVal DecimalIn As Variant, Optional NumberOfBits As Variant) As String
    Dec2Bin = ""
    DecimalIn = Int(CDec(DecimalIn))
    Do While DecimalIn <> 0
        Dec2Bin = Format$(DecimalIn - 2 * Int(DecimalIn / 2)) & Dec2Bin
        DecimalIn = Int(DecimalIn / 2)
    Loop
    If Not IsMissing(NumberOfBits) Then
       If Len(Dec2Bin) > NumberOfBits Then
          Dec2Bin = "Error - Number exceeds specified bit size"
       Else
          Dec2Bin = Right$(String$(NumberOfBits, _
                    "0") & Dec2Bin, NumberOfBits)
       End If
    End If
End Function
 
'Binary To Decimal
' =================
Private Function Bin2Dec(BinaryString As String) As Variant
    Dim X As Integer
    For X = 0 To Len(BinaryString) - 1
        Bin2Dec = CDec(Bin2Dec) + Val(Mid(BinaryString, _
                  Len(BinaryString) - X, 1)) * 2 ^ X
    Next
End Function
Function Ip2Bin(IP As Variant) As String
  IPOctets = Split(IP, ".")
  BinIPOctets = Array(Dec2Bin(IPOctets(0), 8), Dec2Bin(IPOctets(1), 8), Dec2Bin(IPOctets(2), 8), Dec2Bin(IPOctets(3), 8))
  Ip2Bin = Join(BinIPOctets, "")
End Function
Function Bin2Ip(BinIP As Variant) As Variant
    IPOctets = Array(Bin2Dec(Left(BinIP, 8)), Bin2Dec(Mid(BinIP, 9, 8)), Bin2Dec(Mid(BinIP, 17, 8)), Bin2Dec(Right(BinIP, 8)))
    Bin2Ip = Join(IPOctets, ".")
End Function

Function isValidIP(IP As Variant) As Boolean
  isValidIP = True
  IPOctets = Split(IP, ".")
  If (UBound(IPOctets) - LBound(IPOctets) + 1 = 4) Then
    For i = LBound(IPOctets) To UBound(IPOctets)
      If IPOctets(i) > 255 Or IPOctets(i) < 0 Then
        isValidIP = False
      End If
    Next i
  Else
    isValidIP = False
  End If
End Function
Function isValidMask(Mask As Variant) As Variant
  isValidMask = True
  If isValidIP(Mask) Then
    IPMaskOctets = Split(Mask, ".")
    BinIPMaskOctets = Array(Dec2Bin(IPMaskOctets(0), 8), Dec2Bin(IPMaskOctets(1), 8), Dec2Bin(IPMaskOctets(2), 8), Dec2Bin(IPMaskOctets(3), 8))
    BinIPMask = Join(BinIPMaskOctets, "")
    For i = InStr(BinIPMask, "0") To Len(BinIPMask)
      bit = Mid(BinIPMask, i, 1)
      If bit = 1 Then
        isValidMask = False
      End If
    Next i
  Else
    isValidMask = False
  End If
End Function
Function isValidCIDR(CIDR As Variant) As Variant
   isValidCIDR = True
   ip_and_mask = Split(CIDR, "/")
   If (UBound(ip_and_mask) - LBound(ip_and_mask) + 1 = 2) Then
     isValidCIDR = isValidIP(ip_and_mask(LBound(ip_and_mask)))
     Mask = ip_and_mask(UBound(ip_and_mask))
     If (Mask < 0 Or Mask > 32) Then
       isValidCIDR = False
     End If
   Else
     isValidCIDR = False
   End If
End Function
Function ReverseIP(IP As Variant) As Variant
  If isValidIP(IP) Then
    Octets = Split(IP, ".")
    ReverseOctets = Array(Octets(3), Octets(2), Octets(1), Octets(0))
    ReverseIP = Join(ReverseOctets, ".")
  Else
    ReverseIP = "Not valid IP"
  End If
End Function
Function networkIP(IP As Variant, Mask As Variant) As Variant
  If isValidIP(IP) And isValidMask(Mask) Then
    BinIP = Ip2Bin(IP)
    BinMask = Ip2Bin(Mask)
    MaskLen = InStr(BinMask, "0") - 1
    BinIP = Left(BinIP, MaskLen)
    BinIP = Left$(BinIP & String$(32 - MaskLen, "0"), 32)
    networkIP = Bin2Ip(BinIP)
  Else
    errMsg = ""
    If Not isValidIP(IP) Then
      errMsg = "Not valid IP"
    End If
    If Not isValidMask(Mask) Then
      If errMsg <> "" Then
        errMsg = errMsg + " and " + "Not valid Mask"
      Else
        errMsg = "Not valid Mask"
      End If
    End If
    networkIP = errMsg
  End If
End Function
Function broadcastIP(IP As Variant, Mask As Variant) As Variant
  If isValidIP(IP) And isValidMask(Mask) Then
    BinIP = Ip2Bin(IP)
    BinMask = Ip2Bin(Mask)
    MaskLen = InStr(BinMask, "0") - 1
    BinIP = Left(BinIP, MaskLen)
    BinIP = Left$(BinIP & String$(32 - MaskLen, "1"), 32)
    broadcastIP = Bin2Ip(BinIP)
  Else
    errMsg = ""
    If Not isValidIP(IP) Then
      errMsg = "Not valid IP"
    End If
    If Not isValidMask(Mask) Then
      If errMsg <> "" Then
        errMsg = errMsg + " and " + "Not valid Mask"
      Else
        errMsg = "Not valid Mask"
      End If
    End If
    broadcastIP = errMsg
  End If
End Function
Function CIDR(IP As Variant, Mask As Variant) As Variant
  If isValidIP(IP) And isValidMask(Mask) Then
    BinMask = Ip2Bin(Mask)
    MaskLen = InStr(BinMask, "0") - 1
    CIDR = IP & "/" & MaskLen
  Else
    errMsg = ""
    If Not isValidIP(IP) Then
      errMsg = "Not valid IP"
    End If
    If Not isValidMask(Mask) Then
      If errMsg <> "" Then
        errMsg = errMsg + " and " + "Not valid Mask"
      Else
        errMsg = "Not valid Mask"
      End If
    End If
    CIDR = errMsg
  End If
End Function
Function MinHost(IP As Variant, Mask As Variant) As Variant
    netIP = networkIP(IP, Mask)
    BinNetIP = Ip2Bin(netIP)
    BinMinHost = Left(BinNetIP, 31) & "1"
    MinHost = Bin2Ip(BinMinHost)
End Function
Function MaxHost(IP As Variant, Mask As Variant) As Variant
    bcastIP = broadcastIP(IP, Mask)
    BinBroadcastIP = Ip2Bin(bcastIP)
    BinMaxHost = Left(BinBroadcastIP, 31) & "0"
    MaxHost = Bin2Ip(BinMaxHost)
End Function
Function Hostname(FQDN As Variant) As Variant
    Hostname = Left(FQDN, InStr(FQDN, ".") - 1)
End Function
Function Domain(FQDN As Variant) As Variant
    Domain = Right(FQDN, (Len(FQDN) - InStr(FQDN, ".")))
End Function
Function CIDR2Mask(CIDR As Variant) As Variant
    If (Not isValidCIDR(CIDR)) And ((CIDR < 0) Or (CIDR > 32)) Then
      CIDR2Mask = "Not Valid CIDR Provide IP/Mask Ex: (10.0.0.1/24)"
    Else
      If isValidCIDR(CIDR) Then
        ip_and_mask = Split(CIDR, "/")
        MaskLen = ip_and_mask(UBound(ip_and_mask))
      Else
        MaskLen = CIDR
      End If
      BinMask = String(MaskLen, "1") & String(32 - MaskLen, "0")
      CIDR2Mask = Bin2Ip(BinMask)
    End If
End Function
