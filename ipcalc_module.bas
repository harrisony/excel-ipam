Attribute VB_Name = "Module1"
'    Visual Basic for Excel
'    uncomment next line for LibreOffice Basic

'Option VBASupport 1

'    Copyright 2010-2023 Thomas Rohmer-Kretz

'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.

'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.

'    http://trk.free.fr/ipcalc/

'    2023-06-25

Public Const strHex2bin = "0000000100100011010001010110011110001001101010111100110111101111"
Public Const strDec2hex = "000102030405060708090A0B0C0D0E0F101112131415161718191A1B1C1D1E1F202122232425262728292A2B2C2D2E2F303132333435363738393A3B3C3D3E3F404142434445464748494A4B4C4D4E4F505152535455565758595A5B5C5D5E5F606162636465666768696A6B6C6D6E6F707172737475767778797A7B7C7D7E7F808182838485868788898A8B8C8D8E8F909192939495969798999A9B9C9D9E9FA0A1A2A3A4A5A6A7A8A9AAABACADAEAFB0B1B2B3B4B5B6B7B8B9BABBBCBDBEBFC0C1C2C3C4C5C6C7C8C9CACBCCCDCECFD0D1D2D3D4D5D6D7D8D9DADBDCDDDEDFE0E1E2E3E4E5E6E7E8E9EAEBECEDEEEFF0F1F2F3F4F5F6F7F8F9FAFBFCFDFEFF"


'==============================================
'   IP v4
'==============================================

'----------------------------------------------
'   IpIsValid
'----------------------------------------------
' Returns true if an ip address is formated exactly as it should be:
' no space, no extra zero, no incorrect value
Function IpIsValid(ByVal ip As String) As Boolean
    IpIsValid = (IpBinToStr(IpStrToBin(ip)) = ip)
End Function

'----------------------------------------------
'   IpIsPrivate
'----------------------------------------------
' returns TRUE if "ip" is in one of the private IP address ranges
' example:
'   IpIsPrivate("192.168.1.35") returns TRUE
'   IpIsPrivate("209.85.148.104") returns FALSE
Function IpIsPrivate(ByVal ip As String) As Boolean
    IpIsPrivate = (IpIsInSubnet(ip, "10.0.0.0/8") Or IpIsInSubnet(ip, "172.16.0.0/12") Or IpIsInSubnet(ip, "192.168.0.0/16"))
End Function

'----------------------------------------------
'   IpStrToBin
'----------------------------------------------
' Converts a text IP address to binary
' example:
'   IpStrToBin("1.2.3.4") returns 16909060
Function IpStrToBin(ByVal ip As String) As Double
    Dim pos As Integer
    ip = ip + "."
    IpStrToBin = 0
    While ip <> ""
        pos = InStr(ip, ".")
        IpStrToBin = IpStrToBin * 256 + Val(Left(ip, pos - 1))
        ip = Mid(ip, pos + 1)
    Wend
End Function

'----------------------------------------------
'   IpBinToStr
'----------------------------------------------
' Converts a binary IP address to text
' example:
'   IpBinToStr(16909060) returns "1.2.3.4"
Function IpBinToStr(ByVal ip As Double) As String
    Dim divEnt As Double
    Dim i As Integer
    i = 0
    IpBinToStr = ""
    While i < 4
        If IpBinToStr <> "" Then IpBinToStr = "." + IpBinToStr
        divEnt = Int(ip / 256)
        IpBinToStr = Format(ip - (divEnt * 256)) + IpBinToStr
        ip = divEnt
        i = i + 1
    Wend
End Function

'----------------------------------------------
'   IpSubnetToBin
'----------------------------------------------
' Converts a subnet to binary
' This function is similar to IpStrToBin but ignores the host part of the address
' example:
'   IpSubnetToBin("1.2.3.4/24") returns 16909056
'   IpSubnetToBin("1.2.3.0/24") returns 16909056
Function IpSubnetToBin(ByVal ip As String) As Double
    Dim l As Integer
    Dim pos As Integer
    Dim v As Integer
    l = IpSubnetParse(ip)
    ip = ip + "."
    IpSubnetToBin = 0
    While ip <> ""
        pos = InStr(ip, ".")
        v = Val(Left(ip, pos - 1))
        If (l <= 0) Then
            v = 0
        ElseIf (l < 8) Then
            v = v And ((2 ^ l - 1) * 2 ^ (8 - l))
        End If
        IpSubnetToBin = IpSubnetToBin * 256 + v
        ip = Mid(ip, pos + 1)
        l = l - 8
    Wend
End Function

'----------------------------------------------
'   IpAdd
'----------------------------------------------
' example:
'   IpAdd("192.168.1.1"; 4) returns "192.168.1.5"
'   IpAdd("192.168.1.1"; 256) returns "192.168.2.1"
Function IpAdd(ByVal ip As String, offset As Double) As String
    IpAdd = IpBinToStr(IpStrToBin(ip) + offset)
End Function

'----------------------------------------------
'   IpAnd
'----------------------------------------------
' bitwise AND
' example:
'   IpAnd("192.168.1.1"; "255.255.255.0") returns "192.168.1.0"
Function IpAnd(ByVal ip1 As String, ByVal ip2 As String) As String
    ' compute bitwise AND from right to left
    Dim result As String
    While ((ip1 <> "") And (ip2 <> ""))
        Call IpBuild(IpParse(ip1) And IpParse(ip2), result)
    Wend
    IpAnd = result
End Function

'----------------------------------------------
'   IpOr
'----------------------------------------------
' bitwise OR
' example:
'   IpOr("192.168.1.1"; "0.0.0.255") returns "192.168.1.255"
Function IpOr(ByVal ip1 As String, ByVal ip2 As String) As String
    ' compute bitwise OR from right to left
    Dim result As String
    While ((ip1 <> "") And (ip2 <> ""))
        Call IpBuild(IpParse(ip1) Or IpParse(ip2), result)
    Wend
    IpOr = result
End Function

'----------------------------------------------
'   IpXor
'----------------------------------------------
' bitwise XOR
' example:
'   IpXor("192.168.1.1"; "0.0.0.255") returns "192.168.1.254"
Function IpXor(ByVal ip1 As String, ByVal ip2 As String) As String
    ' compute bitwise XOR from right to left
    Dim result As String
    While ((ip1 <> "") And (ip2 <> ""))
        Call IpBuild(IpParse(ip1) Xor IpParse(ip2), result)
    Wend
    IpXor = result
End Function

'----------------------------------------------
'   IpAdd2
'----------------------------------------------
' another implementation of IpAdd which not use the binary representation
Function IpAdd2(ByVal ip As String, offset As Double) As String
    Dim result As String
    While (ip <> "")
        offset = IpBuild(IpParse(ip) + offset, result)
    Wend
    IpAdd2 = result
End Function

'----------------------------------------------
'   IpComp
'----------------------------------------------
' Compares the first 'n' bits of ip1 and ip2
' example:
'   IpComp("10.0.0.0", "10.1.0.0", 9) returns TRUE
'   IpComp("10.0.0.0", "10.1.0.0", 16) returns FALSE
Function IpComp(ByVal ip1 As String, ByVal ip2 As String, ByVal n As Integer) As Boolean
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim mask As Integer
    ip1 = ip1 + "."
    ip2 = ip2 + "."
    While (n > 0) And (ip1 <> "") And (ip2 <> "")
        pos1 = InStr(ip1, ".")
        pos2 = InStr(ip2, ".")
        If n >= 8 Then
            If pos1 <> pos2 Then
                IpComp = False
                Exit Function
            End If
            If Left(ip1, pos1) <> Left(ip2, pos2) Then
                IpComp = False
                Exit Function
            End If
        Else
            mask = (2 ^ n - 1) * 2 ^ (8 - n)
            IpComp = ((Val(Left(ip1, pos1 - 1)) And mask) = (Val(Left(ip2, pos2 - 1)) And mask))
            Exit Function
        End If
        n = n - 8
        ip1 = Mid(ip1, pos1 + 1)
        ip2 = Mid(ip2, pos2 + 1)
    Wend
    IpComp = True
End Function

'----------------------------------------------
'   IpGetByte
'----------------------------------------------
' get one byte from an ip address given its position
' example:
'   IpGetByte("192.168.1.1"; 1) returns 192
Function IpGetByte(ByVal ip As String, pos As Integer) As Integer
    pos = 4 - pos
    For i = 0 To pos
        IpGetByte = IpParse(ip)
    Next
End Function

'----------------------------------------------
'   IpSetByte
'----------------------------------------------
' set one byte in an ip address given its position and value
' example:
'   IpSetByte("192.168.1.1"; 4; 20) returns "192.168.1.20"
Function IpSetByte(ByVal ip As String, pos As Integer, newvalue As Integer) As String
    Dim result As String
    Dim byteval As Double
    i = 4
    While (ip <> "")
        byteval = IpParse(ip)
        If (i = pos) Then byteval = newvalue
        Call IpBuild(byteval, result)
        i = i - 1
    Wend
    IpSetByte = result
End Function

'----------------------------------------------
'   IpMask
'----------------------------------------------
' returns an IP netmask from a subnet
' both notations are accepted
' example:
'   IpMask("192.168.1.1/24") returns "255.255.255.0"
'   IpMask("192.168.1.1 255.255.255.0") returns "255.255.255.0"
Function IpMask(ByVal ip As String) As String
    IpMask = IpBinToStr(IpMaskBin(ip))
End Function

'----------------------------------------------
'   IpWildMask
'----------------------------------------------
' returns an IP Wildcard (inverse) mask from a subnet
' both notations are accepted
' example:
'   IpWildMask("192.168.1.1/24") returns "0.0.0.255"
'   IpWildMask("192.168.1.1 255.255.255.0") returns "0.0.0.255"
Function IpWildMask(ByVal ip As String) As String
    IpWildMask = IpBinToStr(((2 ^ 32) - 1) - IpMaskBin(ip))
End Function

'----------------------------------------------
'   IpInvertMask
'----------------------------------------------
' returns an IP Wildcard (inverse) mask from a subnet mask
' or a subnet mask from a wildcard mask
' example:
'   IpInvertMask("255.255.255.0") returns "0.0.0.255"
'   IpInvertMask("0.0.0.255") returns "255.255.255.0"
Function IpInvertMask(ByVal mask As String) As String
    IpInvertMask = IpBinToStr(((2 ^ 32) - 1) - IpStrToBin(mask))
End Function

'----------------------------------------------
'   IpMaskLen
'----------------------------------------------
' returns prefix length from a mask given by a string notation (xx.xx.xx.xx)
' example:
'   IpMaskLen("255.255.255.0") returns 24 which is the number of bits of the subnetwork prefix
Function IpMaskLen(ByVal ipmaskstr As String) As Integer
    Dim notMask As Double
    notMask = 2 ^ 32 - 1 - IpStrToBin(ipmaskstr)
    zeroBits = 0
    Do While notMask > 0
        notMask = Int(notMask / 2)
        zeroBits = zeroBits + 1
    Loop
    IpMaskLen = 32 - zeroBits
End Function

'----------------------------------------------
'   IpWithoutMask
'----------------------------------------------
' removes the netmask notation at the end of the IP v4 or v6
' example:
'   IpWithoutMask("192.168.1.0/24") returns "192.168.1.1"
'   IpWithoutMask("192.168.1.0 255.255.255.0") returns "192.168.1.1"
'   IpWithoutMask("2001:db8:1:1a0::/59") returns "2001:db8:1:1a0::"
Function IpWithoutMask(ByVal ip As String) As String
    Dim p As Integer
    p = InStr(ip, "/")
    If (p = 0) Then
        p = InStr(ip, " ")
    End If
    If (p = 0) Then
        IpWithoutMask = ip
    Else
        IpWithoutMask = Left(ip, p - 1)
    End If
End Function

'----------------------------------------------
'   IpSubnetLen
'----------------------------------------------
' get the mask len from a subnet
' example:
'   IpSubnetLen("192.168.1.1/24") returns 24
'   IpSubnetLen("192.168.1.1 255.255.255.0") returns 24
Function IpSubnetLen(ByVal ip As String) As Integer
    Dim p As Integer
    p = InStr(ip, "/")
    If (p = 0) Then
        p = InStr(ip, " ")
        If (p = 0) Then
            IpSubnetLen = 32
        Else
            IpSubnetLen = IpMaskLen(Mid(ip, p + 1))
        End If
    Else
        IpSubnetLen = Val(Mid(ip, p + 1))
    End If
End Function

'----------------------------------------------
'   IpSubnetSize
'----------------------------------------------
' returns the number of addresses in a subnet
' example:
'   IpSubnetSize("192.168.1.32/29") returns 8
'   IpSubnetSize("192.168.1.0 255.255.255.0") returns 256
Function IpSubnetSize(ByVal subnet As String) As Double
    IpSubnetSize = 2 ^ (32 - IpSubnetLen(subnet))
End Function

'----------------------------------------------
'   IpClearHostBits
'----------------------------------------------
' set to zero the bits in the host part of an address
' example:
'   IpClearHostBits("192.168.1.1/24") returns "192.168.1.0/24"
'   IpClearHostBits("192.168.1.193 255.255.255.128") returns "192.168.1.128 255.255.255.128"
Function IpClearHostBits(ByVal net As String) As String
    Dim ip As String
    ip = IpWithoutMask(net)
    IpClearHostBits = IpAnd(ip, IpMask(net)) + Mid(net, Len(ip) + 1)
End Function

'----------------------------------------------
'   IpIsInSubnet
'----------------------------------------------
' Returns TRUE if "ip" is in "subnet"
' example:
'   IpIsInSubnet("192.168.1.35"; "192.168.1.32/29") returns TRUE
'   IpIsInSubnet("192.168.1.35"; "192.168.1.32 255.255.255.248") returns TRUE
'   IpIsInSubnet("192.168.1.41"; "192.168.1.32/29") returns FALSE
Function IpIsInSubnet(ByVal ip As String, ByVal subnet As String) As Boolean
    Dim l As Integer
    l = IpSubnetParse(subnet)
    IpIsInSubnet = IpComp(ip, subnet, l)
End Function

'----------------------------------------------
'   IpSubnetMatch
'----------------------------------------------
' Tries to match an IP address or a subnet against a list of subnets in the
' left-most column of table_array and returns the row number
' 'ip' is the value to search for in the subnets in the first column of
'      the table_array
' 'table_array' is one or more columns of data
' 'fast' indicates the search mode : BestMatch or Fast mode
' fast = 0 (default value)
'    This will work on any subnet list. If the search value matches more
'    than one subnet, the smallest subnet will be returned (best match)
' fast = 1
'    The subnet list MUST be sorted in ascending order and MUST NOT contain
'    overlapping subnets. This mode performs a dichotomic search and runs
'    much faster with large subnet lists.
' The function returns 0 if the IP address is not matched.
Function IpSubnetMatch(ByVal ip As String, table_array As Range, Optional fast As Boolean = False) As Long
    Dim i As Long
    IpSubnetMatch = 0
    If fast Then
        Dim a As Long
        Dim b As Long
        Dim ip_bin As Double
        a = 1
        b = table_array.Rows.Count
        ip_bin = IpSubnetToBin(ip)
        Do
            i = (a + b + 0.5) / 2
            If ip_bin < IpSubnetToBin(table_array.Cells(i, 1)) Then
                b = i - 1
            Else
                a = i
            End If
        Loop While a < b
        If IpSubnetIsInSubnet(ip, table_array.Cells(a, 1)) Then
            IpSubnetMatch = a
        End If
    Else
        Dim previousMatchLen As Integer
        Dim searchLen As Integer
        Dim subnet As String
        Dim subnetLen As Integer
        searchLen = IpSubnetParse(ip)
        previousMatchLen = -1
        For i = 1 To table_array.Rows.Count
            subnet = table_array.Cells(i, 1)
            subnetLen = IpSubnetParse(subnet)
            If subnetLen > previousMatchLen Then
                If searchLen >= subnetLen Then
                    If IpComp(ip, subnet, subnetLen) Then
                        previousMatchLen = subnetLen
                        IpSubnetMatch = i
                    End If
                End If
            End If
        Next i
    End If
End Function

'----------------------------------------------
'   IpSubnetVLookup
'----------------------------------------------
' Tries to match an IP address or a subnet against a list of subnets in the
' left-most column of table_array and returns the value in the same row based
' on the index_number
' 'ip' is the value to search for in the subnets in the first column of
'      the table_array
' 'table_array' is one or more columns of data
' 'index_number' is the column number in table_array from which the matching
'      value must be returned. The first column which contains subnets is 1.
' 'fast' indicates the search mode : BestMatch or Fast mode
' fast = 0 (default value)
'    This will work on any subnet list. If the search value matches more
'    than one subnet, the smallest subnet will be returned (best match)
' fast = 1
'    The subnet list MUST be sorted in ascending order and MUST NOT contain
'    overlapping subnets. This mode performs a dichotomic search and runs
'    much faster with large subnet lists.
' Note: add 0.0.0.0/0 in the array if you want the function to return a
' default value (best match mode only)
Function IpSubnetVLookup(ByVal ip As String, table_array As Range, index_number As Integer, Optional fast As Boolean = False) As String
    Dim i As Long
    i = IpSubnetMatch(ip, table_array, fast)
    If i = 0 Then
        IpSubnetVLookup = "Not Found"
    Else
        IpSubnetVLookup = table_array.Cells(i, index_number)
    End If
End Function

'----------------------------------------------
'   IpSubnetVLookupAreas
'----------------------------------------------
' Same as IpSubnetVLookup except that table_array parameter can be a
' named area containing multiple tables. Use it if you want to search in
' more than one table.
' Doesn't have the 'fast' option.
Function IpSubnetVLookupAreas(ByVal ip As String, table_array As Range, index_number As Integer) As String
    Dim i As Long
    Dim previousMatch As String
    previousMatch = "0.0.0.0/0"
    IpSubnetVLookupAreas = "Not Found"
    For a = 1 To table_array.Areas.Count
        For i = 1 To table_array.Areas(a).Rows.Count
            Dim subnet As String
            subnet = table_array.Areas(a).Cells(i, 1)
            If IpIsInSubnet(ip, subnet) And (IpSubnetLen(subnet) > IpSubnetLen(previousMatch)) Then
                previousMatch = subnet
                IpSubnetVLookupAreas = table_array.Areas(a).Cells(i, index_number)
            End If
        Next i
    Next a
End Function

'----------------------------------------------
'   IpSubnetIsInSubnet
'----------------------------------------------
' Returns TRUE if "subnet1" is in "subnet2"
' example:
'   IpSubnetIsInSubnet("192.168.1.35/30"; "192.168.1.32/29") returns TRUE
'   IpSubnetIsInSubnet("192.168.1.41/30"; "192.168.1.32/29") returns FALSE
'   IpSubnetIsInSubnet("192.168.1.35/28"; "192.168.1.32/29") returns FALSE
'   IpSubnetIsInSubnet("192.168.0.128 255.255.255.128"; "192.168.0.0 255.255.255.0") returns TRUE
Function IpSubnetIsInSubnet(ByVal subnet1 As String, ByVal subnet2 As String) As Boolean
    Dim l1 As Integer
    Dim l2 As Integer
    l1 = IpSubnetParse(subnet1)
    l2 = IpSubnetParse(subnet2)
    If l1 < l2 Then
        IpSubnetIsInSubnet = False
    Else
        IpSubnetIsInSubnet = IpComp(subnet1, subnet2, l2)
    End If
End Function

'----------------------------------------------
'   IpDiff
'----------------------------------------------
' difference between 2 IP addresses
' example:
'   IpDiff("192.168.1.7"; "192.168.1.1") returns 6
Function IpDiff(ByVal ip1 As String, ByVal ip2 As String) As Double
    Dim mult As Double
    mult = 1
    IpDiff = 0
    While ((ip1 <> "") Or (ip2 <> ""))
        IpDiff = IpDiff + mult * (IpParse(ip1) - IpParse(ip2))
        mult = mult * 256
    Wend
End Function

'----------------------------------------------
'   IpMaskBin
'----------------------------------------------
' returns binary IP mask from an address with / notation (xx.xx.xx.xx/yy)
' example:
'   IpMask("192.168.1.1/24") returns 4294967040 which is the binary
'   representation of "255.255.255.0"
Function IpMaskBin(ByVal ip As String) As Double
    Dim bits As Integer
    bits = IpSubnetLen(ip)
    IpMaskBin = (2 ^ bits - 1) * 2 ^ (32 - bits)
End Function


'==============================================
'         ARRAY FUNCTIONS
'==============================================
' These functions should be called from spreadsheet cells and return arrays.
' In office 365, call the function from one cell, the resulting array
' will be written in the cell and 'spill' on the cells below.
' In older versions of Excel or in LibreOffice, these functions must
' be called from an array formula.
' How to enter an array formula?
'   select a range of empty cells
'   type the formula
'   press Ctrl+Shift+Enter instead of Enter
'   to modify an array formula, select one of the cells then press Ctrl+/ to
'   automatically select the cell range then modify the formula and press Ctrl+Shift+Enter

'----------------------------------------------
'   IpFindOverlappingSubnets
'----------------------------------------------
' this function will find in the list of subnets which subnets overlap
' 'SubnetsArray' is single column array containing a list of subnets, the
' list may be sorted or not
' the return value is also a array of the same size
' if the subnet on line x is included in a larger subnet from another line,
' this function returns an array in which line x contains the value of the
' larger subnet
' if the subnet on line x is distinct from any other subnet in the array,
' then this function returns on line x an empty cell
' if there are no overlapping subnets in the input array, the returned array
' is empty
Function IpFindOverlappingSubnets(subnets_array As Range) As Variant
    Dim i As Long
    Dim j As Long
    Dim result_array As Variant
    ReDim result_array(1 To subnets_array.Rows.Count, 1 To 1)
    For i = 1 To subnets_array.Rows.Count
        result_array(i, 1) = ""
        For j = 1 To subnets_array.Rows.Count
            If (i <> j) And IpSubnetIsInSubnet(subnets_array.Cells(i, 1), subnets_array.Cells(j, 1)) Then
                result_array(i, 1) = subnets_array.Cells(j, 1)
                Exit For
            End If
        Next j
    Next i
    IpFindOverlappingSubnets = result_array
End Function

'----------------------------------------------
'   IpSortArray
'----------------------------------------------
' 'ip_array' is a single column array containing ip addresses
' the return value is also a array containing the sorted addresses
' 'descending' is an optional parameter, if set to True the result is sorted
' in descending order
Function IpSortArray(ip_array As Range, Optional descending As Boolean = False) As Variant
    Dim list As Variant
    Dim s As Long
    Dim i As Long
    ImportCellRange ip_array, list
    s = UBound(list)
    ' convert IP adresses to sortable numbers
    For i = 1 To s
        list(i) = IpStrToBin(list(i))
    Next i
    
    QuickSort list, 1, s

    ' copy the sorted list as strings in a 2D array
    Dim resultArray As Variant
    ReDim resultArray(1 To s, 1 To 1)
    If descending Then
        For i = 1 To s
            resultArray(s + 1 - i, 1) = IpBinToStr(list(i))
        Next i
    Else
        For i = 1 To s
            resultArray(i, 1) = IpBinToStr(list(i))
        Next i
    End If
    IpSortArray = resultArray
End Function

'----------------------------------------------
'   IpSubnetSortArray
'----------------------------------------------
' 'ip_array' is a single column array containing ip subnets in "prefix/len"
' or "prefix mask" notation
' the return value is also an array of containing the sorted subnets
' 'descending' is an optional parameter, if set to True the result is sorted
' in descending order
Function IpSubnetSortArray(ip_array As Range, Optional descending As Boolean = False) As Variant
    Dim list As Variant
    Dim i As Long
    ImportCellRange ip_array, list
    
    SortRoutes list, False
    
    IpSubnetSortArray = ExportCellRange(list, descending)
End Function

'----------------------------------------------
'   IpSubnetSortJoinArray
'----------------------------------------------
' this function has been removed because the algorithm was faulty for route aggregation
' it has been replaced by IpSubnetAggregateArray and IpRouteAggregateArray

'----------------------------------------------
'   IpSubnetAggregateArray
'----------------------------------------------
' this function can sort and summarize subnets
' 'ip_array' is a single column array containing ip subnets in "prefix/len"
' or "prefix mask" notation, in any order
' the return value is also an array containing a simplified list of subnets:
' - small subnets included in larger subnets are removed
' - contiguous subnets are joined if possible
' - duplicates are removed
' 'descending' is an optional parameter, if set to True the result is sorted
' in descending order
Function IpSubnetAggregateArray(ip_array As Range, Optional descending As Boolean = False) As Variant
    Dim list As Variant
    Dim i As Long
    ImportCellRange ip_array, list

    AggregateSubnets list

    IpSubnetAggregateArray = ExportCellRange(list, descending)
End Function

'----------------------------------------------
'   IpRouteAggregateArray
'----------------------------------------------
' this function can sort and summarize ip routes
' 'ip_array' is a single column array containing ip subnets in "prefix/len next hop"
' or "prefix mask next hop" notation, in any order
' the return value is also an array containing a simplified list of routes:
' - small subnets included in larger subnets are removed
' - contiguous subnets are joined if possible
' - duplicates are removed
' ...except if routes have different next hops and thus can not be summarized
' 'descending' is an optional parameter, if set to True the result is sorted
' in descending order
Function IpRouteAggregateArray(ip_array As Range, Optional descending As Boolean = False) As Variant
    Dim list As Variant
    Dim t As Long
    Dim i As Long
    Dim j As Long
    Dim a As String
    Dim b As String
    Dim len_a As Integer
    Dim len_b As Integer
    Dim nexthop_a As String
    Dim nexthop_b As String
    Dim remove_i As Boolean
    Dim try_join As Boolean
    Dim joined_networks As Boolean

    ImportCellRange ip_array, list
    SortRoutes list, True
    t = UBound(list)

    ' try to join subnets, starting from the end
    i = t
    Do While (i > 1)
        ' first loop from the end, for each subnet 'a' we will try to find if it is
        ' inside a larger subnet with same next hop, or if there is a contiguous subnet of the same size
        ' with the same next hop
        remove_i = False
        try_join = True
        joined_networks = False
        a = IpParseRoute(list(i), nexthop_a)
        len_a = IpSubnetLen(a)
        j = i - 1
        Do While (j >= 1)
            ' loop backward through the routes before 'a'
            b = IpParseRoute(list(j), nexthop_b)
            len_b = IpSubnetLen(b)
            If (len_a > len_b) Then
                try_join = False ' no chance to join 'a' with another network
                If (IpSubnetIsInSubnet(a, b)) Then
                    If (nexthop_a = nexthop_b) Then
                        ' this route is in a larger subnet with same nexthop, it's useless
                        remove_i = True
                    End If
                    ' this route is in a larger subnet with different nexthop, keep it
                    Exit Do
                End If
            ElseIf ((len_a = len_b) And (nexthop_a = nexthop_b) And try_join And (len_b > 0)) Then
                ' b and a may be contiguous, try to create a subnet with a mask 1 bit shorter, see if 'a' fits in it
                bigsubnet = Replace(IpWithoutMask(b) + "/" + Str(len_b - 1), " ", "")
                If (InStr(b, "/") = 0) Then
                    ' change the notation to keep the original "IP mask" notation
                    bigsubnet = IpWithoutMask(b) & " " & IpMask(bigsubnet)
                End If
                If (IpSubnetIsInSubnet(a, bigsubnet)) Then
                    ' OK these subnets can be joined, keep the larger subnet in position j and remove i
                    list(j) = bigsubnet & nexthop_a
                    remove_i = True
                    joined_networks = True
                    Exit Do
                End If
                try_join = False ' no chance to join 'a' with another network before in the list
            End If
            j = j - 1
        Loop
        
        If (remove_i) Then
            ' remove list(i+1) and make the list one element shorter
            For j = i To t - 1
                list(j) = list(j + 1)
            Next j
            t = t - 1
        End If
        If (joined_networks) Then
            ' subnets have been joined and the new subnet may be joined with a network anywhere later in the list, we must restart from the end
            i = t
        Else
            i = i - 1
        End If
    Loop

    ReDim Preserve list(1 To t)
    IpRouteAggregateArray = ExportCellRange(list, descending)
End Function

'----------------------------------------------
'   IpDivideSubnet
'----------------------------------------------
' divide a network in smaller subnets
' "n" is the value that will be added to the subnet length
' "SubnetSeqNbr" is the index of the smaller subnet to return
' example:
'   IpDivideSubnet("1.2.3.0/24"; 2; 0) returns "1.2.3.0/26"
'   IpDivideSubnet("1.2.3.0/24"; 2; 1) returns "1.2.3.64/26"
' if "SubnetSeqNbr" is omitted, this function returns an array with the full list of networks
' example:
'   IpDivideSubnet("1.2.3.0/24"; 2) returns a table of 4 lines containing
'      1.2.3.0/26, 1.2.3.64/26, 1.2.3.128/26, 1.2.3.192/26
'   TRANSPOSE(IpDivideSubnet("1.2.3.0/24"; 2) returns a row of 4 columns
Function IpDivideSubnet(ByVal subnet As String, n As Integer, Optional index As Integer = -1)
    Dim ip As String
    Dim ipbin As Double
    Dim slen As Integer ' length of smaller subnets
    Dim ssize As Long ' size (number of addresses)
    Dim listlen As Integer
    listlen = 2 ^ n ' number of smaller subnets
    ip = IpAnd(IpWithoutMask(subnet), IpMask(subnet))
    ipbin = IpStrToBin(ip)
    slen = IpSubnetLen(subnet) + n
    If (slen > 32) Then
        IpDivideSubnet = "ERR subnet length > 32"
        Exit Function
    End If
    ssize = 2 ^ (32 - slen)
    If (index = -1) Then
        Dim list() As String
        ReDim list(1 To listlen)
        Dim i As Long
        For i = 1 To listlen
            list(i) = Replace(IpBinToStr(ipbin) + "/" + Str(slen), " ", "")
            ipbin = ipbin + ssize
        Next i
        IpDivideSubnet = ExportCellRange(list)
    Else
        If (index >= listlen) Then
            IpDivideSubnet = "ERR index out of range"
            Exit Function
        End If
        ipbin = ipbin + ssize * index
        IpDivideSubnet = Replace(IpBinToStr(ipbin) + "/" + Str(slen), " ", "")
    End If
End Function

'----------------------------------------------
'   IpRangeToCIDR
'----------------------------------------------
' returns a network or a list of networks given the first and the
' last address of an IP range
' if this function is used in a array formula, it may return more
' than one network
' example:
'   IpRangeToCIDR("10.0.0.1","10.0.0.254") returns 10.0.0.0/24
'   IpRangeToCIDR("10.0.0.1","10.0.1.63") returns the array : 10.0.0.0/24 10.0.1.0/26
' note:
'   10.0.0.0 or 10.0.0.1 as the first address returns the same result
'   10.0.0.254 or 10.0.0.255 (broadcast) as the last address returns the same result
Function IpRangeToCIDR(ByVal firstAddr As String, ByVal lastAddr As String) As Variant
    firstAddr = IpAnd(firstAddr, "255.255.255.254") ' set the last bit to zero
    lastAddr = IpOr(lastAddr, "0.0.0.1") ' set the last bit to one
    Dim list() As String
    Dim n As Long
    n = 0
    Do
        l = 0
        Do ' find the largest network which first address is firstAddr and which last address is not higher than lastAddr
            ' build a network of length l
            ' if it does not comply the above conditions, try with a smaller network
            l = l + 1
            net = firstAddr & "/" & l
            ip1 = IpAnd(firstAddr, IpMask(net)) ' first @ of this network
            ip2 = IpOr(firstAddr, IpWildMask(net)) ' last @ of this network
            net = ip1 & "/" & l ' rebuild the network with the first address
            diff = IpDiff(ip2, lastAddr) ' difference between the last @ of this network and the lastAddr we need to reach
        Loop While (l < 32) And ((ip1 <> firstAddr) Or (diff > 0))
        
        n = n + 1
        ReDim Preserve list(1 To n)
        list(n) = net
        firstAddr = IpAdd(ip2, 1)
    Loop While (diff < 0) ' if we haven't reached the lastAddr, loop to build another network
    
    IpRangeToCIDR = ExportCellRange(list)
End Function

'----------------------------------------------
'   IpSubtractSubnets
'----------------------------------------------
' removes subnets from a list of subnets
' this function must be used in an array formula
' 'input_array' is a list of assigned subnets
' 'subtract_array' is a list of used subnets
' the result is a list of unused subnets
' fast=False will preserve the order and structure of the original list as much as possible
' fast=True (default and recommanded) will return a sorted and aggregated list of networks
Function IpSubtractSubnets(input_array As Range, subtract_array As Range, Optional fast As Boolean = True) As Variant
    Dim list As Variant
    Dim subtract As Variant
    ImportCellRange input_array, list
    ImportCellRange subtract_array, subtract

    SubtractSubnets list, subtract, fast

    IpSubtractSubnets = ExportCellRange(list)
End Function

'----------------------------------------------
'   IpCommonSubnets
'----------------------------------------------
' returns the list of IP ranges which are common to both input lists
Function IpCommonSubnets(input_array1 As Range, input_array2 As Range) As Variant
    Dim list1 As Variant
    Dim list2 As Variant
    Dim list As Variant

    ImportCellRange input_array1, list1
    ImportCellRange input_array2, list2

    list = list1
    SubtractSubnets list1, list2
    SubtractSubnets list, list1
    
    IpCommonSubnets = ExportCellRange(list)
End Function

'==============================================
'         INTERNAL FUNCTIONS
'==============================================

'----------------------------------------------
'   IpSubnetParse
'----------------------------------------------
' Get the mask len from a subnet and remove the mask from the address
' The ip parameter is modified and the subnet mask is removed
' example:
'   IpSubnetParse("192.168.1.1/24") returns 24 and ip is changed to "192.168.1.1"
'   IpSubnetParse("192.168.1.1 255.255.255.0") returns 24 and ip is changed to "192.168.1.1"
Function IpSubnetParse(ByRef ip As String) As Integer
    Dim p As Integer
    p = InStr(ip, "/")
    If (p = 0) Then
        p = InStr(ip, " ")
        If (p = 0) Then
            IpSubnetParse = 32
        Else
            IpSubnetParse = IpMaskLen(Mid(ip, p + 1))
            ip = Left(ip, p - 1)
        End If
    Else
        IpSubnetParse = Val(Mid(ip, p + 1))
        ip = Left(ip, p - 1)
    End If
End Function

' this function is used by IpRouteAggregateArray to extract the subnet
' and next hop in route
' the supported formats are
' 10.0.0.0 255.255.255.0 1.2.3.4
' 10.0.0.0/24 1.2.3.4
' the next hop can be any character sequence, and not only an IP
Function IpParseRoute(ByVal route As String, ByRef nexthop As String)
    slash = InStr(route, "/")
    sp = InStr(route, " ")
    If ((slash = 0) And (sp > 0)) Then
        temp = Mid(route, sp + 1)
        sp = InStr(sp + 1, route, " ")
    End If
    If (sp = 0) Then
        IpParseRoute = route
        nexthop = ""
    Else
        IpParseRoute = Left(route, sp - 1)
        nexthop = Mid(route, sp)
    End If
End Function


' internal function
' sort a list of subnets or routes, in place
Sub SortRoutes(list As Variant, clearHostBits As Boolean)
    Dim s As Long
    Dim i As Long
    Dim route As String
    Dim subnet As String
    Dim nexthop As String
    s = UBound(list)
    ' add a sort hex key
    For i = 1 To s
        route = list(i)
        subnet = IpClearHostBits(IpParseRoute(route, nexthop))
        If clearHostBits Then route = subnet & nexthop ' rewrite the route with the modified subnet
        list(i) = IpStrToHex(IpWithoutMask(subnet)) + ByteToHex(IpSubnetLen(subnet)) + route
    Next i
    
    QuickSort list, 1, s

    ' remove the 10-character key
    For i = 1 To s
        list(i) = Mid(list(i), 11)
    Next i
End Sub

' internal function
' sort and aggregate a list of subnets, in place
Sub AggregateSubnets(list As Variant)
    Dim t As Long
    Dim i As Long
    Dim j As Long
    Dim a As String
    Dim b As String
    Dim remove_next As Boolean

    SortRoutes list, True
    t = UBound(list)

    ' try to join subnets
    i = 1
    Do While (i < t)
        remove_next = False
        a = list(i)
        b = list(i + 1)
        If (IpSubnetIsInSubnet(b, a)) Then
            ' b is in a or b == a
            remove_next = True
        ElseIf (IpSubnetLen(a) = IpSubnetLen(b)) Then
            ' a and b may be contiguous, try to create a subnet with a mask 1 bit short, see if 'b' fits in it
            bigsubnet = Replace(IpWithoutMask(a) + "/" + Str(IpSubnetLen(a) - 1), " ", "")
            If (InStr(a, "/") = 0) Then
                ' change the notation to keep the original "IP mask" notation
                bigsubnet = IpWithoutMask(a) & " " & IpMask(bigsubnet)
            End If
            If (IpSubnetIsInSubnet(b, bigsubnet)) Then
                ' OK these subnets can be joined, keep the larger subnet in position i and remove next line
                list(i) = bigsubnet
                remove_next = True
            End If
        End If
        
        If (remove_next) Then
            ' remove list(i+1) and make the list one element shorter
            For j = i + 1 To t - 1
                list(j) = list(j + 1)
            Next j
            t = t - 1
            ' step back and try again because list(i) may be joined with list(i-1)
            If (i > 1) Then i = i - 1
        Else
            i = i + 1
        End If
    Loop
    ReDim Preserve list(1 To t)
End Sub

' internal function
' remove subnets from a list of nets
Sub SubtractSubnets(list As Variant, subtract_array As Variant, Optional fast As Boolean = True)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim s As Long
    Dim net As String
    Dim subtractNet As String
    
    If fast Then
        AggregateSubnets list
        AggregateSubnets subtract_array
    End If

    s = UBound(list)

    i = 1
    j = 1
    Do
        subtractNet = subtract_array(i)
        net = list(j)
        ' is the network to remove equal or larger ?
        If IpSubnetIsInSubnet(net, subtractNet) Then ' remove the network from input_array
            For k = j To s - 1
                list(k) = list(k + 1)
            Next k
            s = s - 1
        ' is the network to remove smaller ?
        ElseIf IpSubnetIsInSubnet(subtractNet, net) Then ' split this network in input_array
            ' insert a line in the result array
            s = s + 1
            If s > UBound(list) Then
                ReDim Preserve list(1 To s)
            End If
            For k = s To j + 2 Step -1
                list(k) = list(k - 1)
            Next k
            ' create 2 smaller subnets
            list(j + 1) = IpDivideSubnet(list(j), 1, 1)
            list(j) = IpDivideSubnet(list(j), 1, 0)
            ' nothing more to do here, on next loop we will run through these 2 new networks
            ' and we will match or continue to divide one of them
        Else
            ' nothing to do, the networks have no address in common
            If fast Then
                ' both lists are ordered so we can iterate in both lists at the same time
                ' with only one loop, wich is much faster with large lists
                ' the network with the lower address is skiped
                If IpSubnetToBin(net) < IpSubnetToBin(subtractNet) Then
                    j = j + 1
                Else
                    i = i + 1
                End If
            Else
                ' the lists are unordered so we must test each element in subtract_array with
                ' each element in list (2 embedded loops)
                j = j + 1
                If j > s Then
                    i = i + 1
                    j = 1
                End If
            End If
        End If
    Loop While (i <= UBound(subtract_array)) And (j <= s)

    ReDim Preserve list(1 To s)
End Sub

'----------------------------------------------
'   IpParse
'----------------------------------------------
' Parses an IP address by iteration from right to left
' Removes one byte from the right of "ip" and returns it as an integer
' example:
'   if ip="192.168.1.32"
'   IpParse(ip) returns 32 and ip="192.168.1" when the function returns
Function IpParse(ByRef ip As String) As Integer
    Dim pos As Integer
    pos = InStrRev(ip, ".")
    If pos = 0 Then
        IpParse = Val(ip)
        ip = ""
    Else
        IpParse = Val(Mid(ip, pos + 1))
        ip = Left(ip, pos - 1)
    End If
End Function

'----------------------------------------------
'   IpBuild
'----------------------------------------------
' Builds an IP address by iteration from right to left
' Adds "ip_byte" to the left the "ip"
' If "ip_byte" is greater than 255, only the lower 8 bits are added to "ip"
' and the remaining bits are returned to be used on the next IpBuild call
' example 1:
'   if ip="168.1.1"
'   IpBuild(192, ip) returns 0 and ip="192.168.1.1"
' example 2:
'   if ip="1"
'   IpBuild(258, ip) returns 1 and ip="2.1"
Function IpBuild(ip_byte As Double, ByRef ip As String) As Double
    If ip <> "" Then ip = "." + ip
    ip = Format(ip_byte And 255) + ip
    IpBuild = ip_byte \ 256
End Function

'==============================================
'   IP v6
'==============================================

'----------------------------------------------
'   Ipv6MaskLen
'----------------------------------------------
' returns prefix length from an IPv6 net
' example:
'   Ipv6MaskLen("2001:db8:1f89::/48") returns 48
Function Ipv6MaskLen(ByVal CIDRNet As String) As Integer
    slash = InStr(CIDRNet, "/")
    If (slash = 0) Then
        Ipv6MaskLen = 128
    Else
        Ipv6MaskLen = Val(Mid(CIDRNet, slash + 1))
    End If
End Function

'----------------------------------------------
'   Ipv6WithoutMask
'----------------------------------------------
' removes the /xx netmask notation at the end of the IP
' example:
'   Ipv6WithoutMask("2001:db8:1f89::/48") returns "2001:db8:1f89::"
Function Ipv6WithoutMask(ByVal CIDRNet As String) As String
    slash = InStr(CIDRNet, "/")
    If (slash = 0) Then
        Ipv6WithoutMask = CIDRNet
    Else
        Ipv6WithoutMask = Left(CIDRNet, slash - 1)
    End If
End Function

'----------------------------------------------
'   Ipv6IsInSubnet
'----------------------------------------------
' returns TRUE if "ip" is in "subnet"
' example:
'   Ipv6IsInSubnet("2001:db8:1:::ac1f:1"; "2001:db8:1::/48") returns TRUE
'   Ipv6IsInSubnet("2001:db8:2:::ac1f:1"; "2001:db8:1::/48") returns FALSE
Function Ipv6IsInSubnet(ByVal ip As String, ByVal subnet As String) As Variant
    prefixlen = Ipv6MaskLen(subnet)
    subnet = Ipv6ToBin(subnet)
    ip = Ipv6ToBin(ip)
    If (Left(subnet, prefixlen) = Left(ip, prefixlen)) Then
        Ipv6IsInSubnet = True
    Else
        Ipv6IsInSubnet = False
    End If
End Function

'----------------------------------------------
'   Ipv6SubnetIsInSubnet
'----------------------------------------------
' Returns TRUE if "subnet1" is in "subnet2"
' example:
Function Ipv6SubnetIsInSubnet(ByVal subnet1 As String, ByVal subnet2 As String) As Boolean
    Dim l1 As Integer
    Dim l2 As Integer
    l1 = Ipv6SubnetParse(subnet1)
    l2 = Ipv6SubnetParse(subnet2)
    If l1 < l2 Then
        Ipv6SubnetIsInSubnet = False
    Else
        Ipv6SubnetIsInSubnet = (Left(Ipv6ToBin(subnet1), l2) = Left(Ipv6ToBin(subnet2), l2))
    End If
End Function

'----------------------------------------------
'   Ipv6IsValid
'----------------------------------------------
' Returns true if an ipv6 address has a valid format:
' no space, no more than 4 digits by group, correct number of : and ::
Function Ipv6IsValid(ByVal ip As String) As Variant
    d = 0 ' number of double columns
    c = 0 ' number of columns
    n = 0 ' number of digits by group
    Ipv6IsValid = False
    For i = 1 To Len(ip)
        If (Mid(ip, i, 2) = "::") Then
            d = d + 1
            c = c + 1
            n = 0
        ElseIf (Mid(ip, i, 1) = ":") Then
            c = c + 1
            n = 0
        ElseIf (InStr("0123456789abcdefABCDEF", Mid(ip, i, 1)) > 0) Then
            n = n + 1
            If (n > 4) Then Exit Function ' to many digits in block
        Else
            Exit Function ' invalid character
        End If
    Next
    If ((d = 0) And (c = 7)) Then
        Ipv6IsValid = True
    ElseIf ((d = 1) And (c <= 7)) Then
        Ipv6IsValid = True
    Else
        Ipv6IsValid = False
    End If
End Function

'----------------------------------------------
'   Ipv6AddMissingColumns
'----------------------------------------------
' this function is called from Ipv6Expand and replaces the :: by the
' right amount of :
' examples:
'   Ipv6AddMissingColumns(1:2:3::8) returns "1:2:3:::::8"
'   Ipv6AddMissingColumns(1:2:3:4:5::8) returns "1:2:3:4:5:::8"
'   Ipv6AddMissingColumns(1:2:3::) returns "1:2:3:::::"
Function Ipv6AddMissingColumns(ByVal ip As String) As Variant
    d = 0 ' number of double columns
    c = 0 ' number of columns
    For i = 1 To Len(ip)
        If (Mid(ip, i, 2) = "::") Then d = d + 1
        If (Mid(ip, i, 1) = ":") Then c = c + 1
    Next
    If ((d = 0) And (c = 7)) Then
        ' 7 single columns, nothing to do
        ip2 = ip
    ElseIf (d = 1) Then
        ' one double columns, replace with the right number of columns
        ip2 = Replace(ip, "::", Left("::::::::", 9 - c))
    Else
        ' any other case is an error
        Ipv6AddMissingColumns = "0:0:0:0:0:0:0:0"
        Exit Function
    End If
    Ipv6AddMissingColumns = ip2
End Function

'----------------------------------------------
'   Ipv6Expand
'----------------------------------------------
' returns a representation of an IPv6 address with all the missing zeros
' the result has a fixed length of 39 caracters
' example :
'   Ipv6Expand("1:2:3::8") returns "0001:0002:0003:0000:0000:0000:0000:0008"
Function Ipv6Expand(ByVal ip As String) As Variant
    ip = "0" & Ipv6AddMissingColumns(Ipv6WithoutMask(ip))
    While (ip <> "")
        ip2 = Ipv6Parse(ip) & ip2
        If (ip <> "") Then
            ip2 = ":" & ip2
        End If
    Wend
    Ipv6Expand = ip2
End Function

'----------------------------------------------
'   Ipv6Compress
'----------------------------------------------
' returns the shortest representation of an IPv6 address
' examples:
'   Ipv6Compress("0001:0002:0003:0000:0000:0000:0000:0008") returns "1:2:3::8"
'   Ipv6Compress("01:0:0::") returns "1::"
Function Ipv6Compress(ByVal ip As String) As String
    Dim ip2 As String, ip3 As String, ip4 As String
    
    ' start with the expanded representation of ip
    ip2 = Ipv6Expand(ip)
    ' rebuild ip, this will remove zeros at the begining of each hex block
    ' if a block is null, this will keep one zero
    While (ip2 <> "")
        offset = Ipv6Build(Ipv6ParseInt(ip2), ip3)
    Wend

    ' try to replace the longuest sequence of zero blocks by ::
    s = ":0:0:0:0:0:0:"
    For i = Len(s) To 3 Step -2
        ip4 = Replace(ip3, Left(s, i), "::", 1, 1)
        If (ip3 <> ip4) Then Exit For
    Next
    
    ' remove first 0 if ip starts with 0::
    If (Left(ip4, 3) = "0::") Then ip4 = Mid(ip4, 2)
    ' remove last 0 if ip ends with ::0
    If (Right(ip4, 3) = "::0") Then ip4 = Left(ip4, Len(ip4) - 1)

    Ipv6Compress = ip4
End Function

'----------------------------------------------
'   Ipv6ToBin
'----------------------------------------------
' returns a string representing the binary value of IPv6 address
' the result has a fixed length of 128 characters
Function Ipv6ToBin(ByVal ip As String) As Variant
    Dim result As String
    ip2 = Replace(Ipv6Expand(ip), ":", "")
    For i = 1 To Len(ip2)
        Dim code As Integer
        code = Asc(Mid$(ip2, i, 1))
        If code > 96 Then ' "a" to "f" --> 10 to 15
            code = code - 87
        ElseIf code > 64 Then ' "A" to "F" --> 10 to 15
            code = code - 55
        Else ' "0" to "9" --> 0 to 9
            code = code - 48
        End If
        result = result & Mid$(strHex2bin, code * 4 + 1, 4)
    Next
    Ipv6ToBin = result
End Function

'----------------------------------------------
'   Ipv6FromBin
'----------------------------------------------
' returns an IPv6 from a string representing the binary value of IPv6 address
' the parameter must be a 128 character string
Function Ipv6FromBin(ByVal ipbin As String) As Variant
    Dim result As String
    Dim pos As Integer
    pos = 1
    If Len(ipbin) <> 128 Then
        Ipv6FromBin = ""
        Exit Function
    End If
    
    For bloc = 1 To 8
        Dim v As Double
        v = 0
        For bit = 1 To 16
            v = v * 2 + Val(Mid(ipbin, pos, 1))
            pos = pos + 1
        Next
        result = result + LCase(Hex(v))
        If (bloc < 8) Then result = result + ":"
    Next
    Ipv6FromBin = Ipv6Compress(result)
End Function

'----------------------------------------------
'   Ipv6AddInt
'----------------------------------------------
' Add a value to an IPv6 address
' example:
'   Ipv6AddInt("1::2"; 16) returns "1:12"
Function Ipv6AddInt(ByVal ip As String, offset As Double) As String
    Dim result As String
    ip = Ipv6Expand(ip)
    While (ip <> "")
        offset = Ipv6Build(Ipv6ParseInt(ip) + offset, result)
    Wend
    Ipv6AddInt = Ipv6Compress(result)
End Function

'----------------------------------------------
'   Ipv6Add
'----------------------------------------------
' Add two IPv6 addresses
' example:
'   Ipv6Add("1:2::"; "::3") returns "1:2::3"
'   Ipv6Add("1:2::2"; "::3") returns "1:2::5"
Function Ipv6Add(ByVal ip1 As String, ByVal ip2 As String) As String
    Dim result As String
    Dim offset As Double
    ip1 = Ipv6Expand(ip1)
    ip2 = Ipv6Expand(ip2)
    While ((ip1 <> "") And (ip2 <> ""))
        offset = Ipv6Build(Ipv6ParseInt(ip1) + Ipv6ParseInt(ip2) + offset, result)
    Wend
    Ipv6Add = Ipv6Compress(result)
End Function

'----------------------------------------------
'   Ipv6GetBlock
'----------------------------------------------
' Returns the 4-digit hexa block at position blockNbr
' The value of blockNbr can be 1 to 8, block 1 is the block on the left.
' example:
'   Ipv6GetBlock("2001:db8:1f89:c5a3::ac1f:8001"; 2) returns "0db8"
Function Ipv6GetBlock(ByVal ip As String, blockNbr As Integer) As String
    Ipv6GetBlock = Mid(Ipv6Expand(ip), blockNbr * 5 - 4, 4)
End Function

'----------------------------------------------
'   Ipv6GetBlockInt
'----------------------------------------------
' Same as above except that the returned value is an integer between
' 0 and 65535
Function Ipv6GetBlockInt(ByVal ip As String, blockNbr As Integer) As Double
    Ipv6GetBlockInt = Hex2Bin(Ipv6GetBlock(ip, blockNbr))
End Function

'----------------------------------------------
'   Ipv6SetBlock
'----------------------------------------------
' Sets the value of the 4-digit hexa block at position blockNbr
' The value of blockNbr can be 1 to 8, block 1 is the block on the left.
' example:
'   Ipv6SetBlock("2001::"; 2; "db8") returns "2001:0db8::"
Function Ipv6SetBlock(ByVal ip As String, blockNbr As Integer, ByVal valHex As String) As String
    ' make valHex exactly 4 characters long
    valHex = Right("0000" & valHex, 4)
    ip = Ipv6Expand(ip)
    Mid(ip, blockNbr * 5 - 4, 4) = valHex
    Ipv6SetBlock = Ipv6Compress(ip)
End Function

'----------------------------------------------
'   Ipv6SetBlockInt
'----------------------------------------------
' Same as above except that the block value is passed as an integer between
' 0 and 65535
Function Ipv6SetBlockInt(ByVal ip As String, blockNbr As Integer, valInt As Double) As String
    Dim valHex As String
    valHex = LCase(Hex(valInt And 65535))
    Ipv6SetBlockInt = Ipv6SetBlock(ip, blockNbr, valHex)
End Function

'----------------------------------------------
'   Ipv6SetBits
'----------------------------------------------
' Sets one or more bits in a ip v6 addresse
' bits is a string with one or more "0" and "1"
' offset is the position of the first bit to set between 1 to 128 from left to right
Function Ipv6SetBits(ByVal ip As String, bits As String, offset As Integer) As String
    Dim ipbin As String
    Dim result As String
    ipbin = Ipv6ToBin(ip) ' convert to binary
    result = Left(ipbin, offset - 1) + bits
    If Len(result) < 128 Then
        result = result + Right(ipbin, 128 - Len(result))
    End If
    result = Left(result, 128) ' make sure we do not exceed 128 bits
    Ipv6SetBits = Ipv6FromBin(result)
End Function

'----------------------------------------------
'   Ipv6GetIpv4
'----------------------------------------------
' Get the value of an IPv4 in an IPv6 at a given position
' exemple:
'    Ipv6GetIpv4("2001:c0a8:102::"; 2) returns "192.168.1.2"
Function Ipv6GetIpv4(ByVal ipv6 As String, blockNbr As Integer) As String
    Ipv6GetIpv4 = IpBinToStr(Ipv6GetBlockInt(ipv6, blockNbr) * 65536 + Ipv6GetBlockInt(ipv6, blockNbr + 1))
End Function

'----------------------------------------------
'   Ipv6SetIpv4
'----------------------------------------------
' Put the value of an IPv4 in an IPv6 at a given position
' exemple:
'    Ipv6SetIpv4("2001::"; 2; "192.168.1.2") returns "2001:c0a8:102::"
Function Ipv6SetIpv4(ByVal ipv6 As String, blockNbr As Integer, ByVal ipv4 As String) As String
    Dim result As String
    
    byte1 = IpParse(ipv4)
    byte2 = IpParse(ipv4)
    byte3 = IpParse(ipv4)
    byte4 = IpParse(ipv4)
    
    result = Ipv6SetBlockInt(ipv6, blockNbr + 1, byte1 + 256 * byte2)
    Ipv6SetIpv4 = Ipv6SetBlockInt(result, blockNbr, byte3 + 256 * byte4)
End Function

'----------------------------------------------
'   Ipv6SubnetFirstAddress
'----------------------------------------------
' example:
'   Ipv6SubnetFirstAddress("2001:db8:1:1a0::/59") returns "2001:db8:1:1a0::"
Function Ipv6SubnetFirstAddress(ByVal subnet As String) As Variant
    prefixlen = Ipv6MaskLen(subnet)
    Ipv6SubnetFirstAddress = Ipv6SetBits(subnet, String(128 - prefixlen, "0"), prefixlen + 1)
End Function

'----------------------------------------------
'   Ipv6SubnetLastAddress
'----------------------------------------------
' example:
'   Ipv6SubnetLastAddress("2001:db8:1:1a0::/59") returns "2001:db8:1:1bf:ffff:ffff:ffff:ffff"
Function Ipv6SubnetLastAddress(ByVal subnet As String) As Variant
    prefixlen = Ipv6MaskLen(subnet)
    Ipv6SubnetLastAddress = Ipv6SetBits(subnet, String(128 - prefixlen, "1"), prefixlen + 1)
End Function

'----------------------------------------------
'   Ipv6Match
'----------------------------------------------
' Tries to match an IP address or a subnet against a list of subnets in the
' left-most column of table_array and returns the row number
' 'ip' is the ip or network to search for in the subnets in the first column of
'      the table_array
' 'table_array' is one or more columns of data
' 'fast' indicates the search mode : BestMatch or Fast mode
' fast = 0 (default value)
'    This will work on any subnet list. If the search value matches more
'    than one subnet, the smallest subnet will be returned (best match)
' fast = 1
'    The subnet list MUST be sorted in ascending order and MUST NOT contain
'    overlapping subnets. This mode performs a dichotomic search and runs
'    much faster with large subnet lists.
' The function returns 0 if the IP address is not matched.
Function Ipv6Match(ByVal ip As String, table_array As Range, Optional fast As Boolean = False) As Long
    Dim i As Long
    Ipv6Match = 0
    If fast Then
        Dim a As Long
        Dim b As Long
        Dim ipexp As String ' expanded ip without mask
        ipexp = Ipv6Expand(ip)
        a = 1
        b = table_array.Rows.Count
        Do
            i = (a + b + 0.5) / 2
            If StrComp(ipexp, Ipv6Expand(table_array.Cells(i, 1)), 1) = -1 Then
                b = i - 1
            Else
                a = i
            End If
        Loop While a < b
        If Ipv6SubnetIsInSubnet(ip, table_array.Cells(a, 1)) Then
            Ipv6Match = a
        End If
    Else
        Dim previousMatchLen As Integer
        Dim searchLen As Integer
        Dim subnet As String
        Dim subnetLen As Integer
        Dim ipbin As String
        searchLen = Ipv6SubnetParse(ip)
        ipbin = Ipv6ToBin(ip)
        previousMatchLen = -1
        For i = 1 To table_array.Rows.Count
            subnet = table_array.Cells(i, 1)
            subnetLen = Ipv6SubnetParse(subnet)
            If subnetLen > previousMatchLen Then
                If searchLen >= subnetLen Then
                    If Left(ipbin, subnetLen) = Left(Ipv6ToBin(subnet), subnetLen) Then
                        previousMatchLen = subnetLen
                        Ipv6Match = i
                    End If
                End If
            End If
        Next i
    End If
End Function

'----------------------------------------------
'   Ipv6VLookup
'----------------------------------------------
' Tries to match an IP address or a subnet against a list of subnets in the
' left-most column of table_array and returns the value in the same row based
' on the index_number
' 'ip' is the value to search for in the subnets in the first column of
'      the table_array
' 'table_array' is one or more columns of data
' 'index_number' is the column number in table_array from which the matching
'      value must be returned. The first column which contains subnets is 1.
' 'fast' indicates the search mode : BestMatch or Fast mode
' fast = 0 (default value)
'    This will work on any subnet list. If the search value matches more
'    than one subnet, the smallest subnet will be returned (best match)
' fast = 1
'    The subnet list MUST be sorted in ascending order and MUST NOT contain
'    overlapping subnets. This mode performs a dichotomic search and runs
'    much faster with large subnet lists.
' Note: add ::/0 in the array if you want the function to return a
' default value (best match mode only)
Function Ipv6VLookup(ByVal ip As String, table_array As Range, index_number As Integer, Optional fast As Boolean = False) As String
    Dim i As Long
    i = Ipv6Match(ip, table_array, fast)
    If i = 0 Then
        Ipv6VLookup = "Not Found"
    Else
        Ipv6VLookup = table_array.Cells(i, index_number)
    End If
End Function

'==============================================
'   IP v6 internal functions
'==============================================

'----------------------------------------------
'   Ipv6Parse
'----------------------------------------------
' Parses an IPv6 address by iteration from right to left
' Removes a 16-bit value from the right of "ip" and returns it as 4 character
' long hex value
' Important: This function does not expand the :: if there is any
' example:
'   if ip="1:2:3:4:5:6:7:8"
'   IpParse(ip) returns "0008" and ip="1:2:3:4:5:6:7" when the function returns
Function Ipv6Parse(ByRef ip As String) As String
    Dim pos As Integer
    pos = InStrRev(ip, ":")
    If pos = 0 Then
        v = ip
        ip = ""
    Else
        v = Mid(ip, pos + 1)
        ip = Left(ip, pos - 1)
    End If
    Ipv6Parse = Right("0000" & v, 4)
End Function

'----------------------------------------------
'   Ipv6ParseInt
'----------------------------------------------
' Same as Ipv6Parse but returns a Double instead of String
Function Ipv6ParseInt(ByRef ip As String) As Double
    Dim v As Double
    strHex = Ipv6Parse(ip)
    For i = 1 To Len(strHex)
        v = 16 * v + Val("&H" & Mid$(strHex, i, 1))
    Next
    Ipv6ParseInt = v
End Function

'----------------------------------------------
'   Ipv6Build
'----------------------------------------------
' Builds an IP address by iteration from right to left
' Adds "v16bits" to the left the "ip"
' If "v16bits" is greater than 65535 (= FFFF), only the lower 16 bits are
' added to "ip" and the remaining bits are returned to be used on the next
' IpBuild call
Function Ipv6Build(v16bits As Double, ByRef ip As String) As Double
    If ip <> "" Then ip = ":" + ip
    ip = LCase(Hex(v16bits And 65535)) + ip
    Ipv6Build = v16bits \ 65536
End Function

'----------------------------------------------
'   Ipv6SubnetParse
'----------------------------------------------
' Get the mask len from a subnet and remove the mask from the address
' The ip parameter is modified and the subnet mask is removed
' example:
'   Ipv6SubnetParse("2001:db8:1:1a0::/59") returns 59 and ip is changed to "2001:db8:1:1a0::"
'   Ipv6SubnetParse("2001:db8:1:1a0::") returns 128 and ip is unchanged
Function Ipv6SubnetParse(ByRef ip As String) As Integer
    Dim p As Integer
    p = InStr(ip, "/")
    If (p = 0) Then
        Ipv6SubnetParse = 128
    Else
        Ipv6SubnetParse = Val(Mid(ip, p + 1))
        ip = Left(ip, p - 1)
    End If
End Function

'==============================================
'   global internal functions
'==============================================

'----------------------------------------------
'   Hex2Bin
'----------------------------------------------
' Converts an hex value to binary representation
' example:
'   Hex2Bin("FF00") returns "1111111100001111"
Function Hex2Bin(ByVal strHex As String) As Double
    Dim v As Double
    For i = 1 To Len(strHex)
        v = 16 * v + Val("&H" & Mid$(strHex, i, 1))
    Next
    Hex2Bin = v
End Function

'----------------------------------------------
'   ByteToHex
'----------------------------------------------
' Converts a byte to a 2-character Hex string
Function ByteToHex(b As Integer) As String
    b = b And 255 ' make sure 0 <= b < 256
    ByteToHex = Mid(strDec2hex, b * 2 + 1, 2)
End Function

'----------------------------------------------
'   IpStrToHex
'----------------------------------------------
' Converts a text IP address to a 8-character Hex string
' example:
'   IpStrToBin("192.168.1.255") returns "C0A801FF"
Function IpStrToHex(ByVal ip As String) As String
    Dim pos As Integer
    Dim res As String
    ip = ip + "."
    res = ""
    While ip <> ""
        pos = InStr(ip, ".")
        res = res + Mid(strDec2hex, Val(Left(ip, pos - 1)) * 2 + 1, 2)
        ip = Mid(ip, pos + 1)
    Wend
    IpStrToHex = Right("00000000" + res, 8)
End Function

'----------------------------------------------
'   ImportCellRange
'----------------------------------------------
' the purpose of this function is to handle Range objects which are passed to functions
' instead of regular arrays when called from a spreadsheet cell
' it also removes empty cells which may be found in the cell range
Sub ImportCellRange(cell_range As Range, list As Variant)
    Dim s As Long
    Dim t As Long
    t = 0
    s = cell_range.Rows.Count
    ReDim list(1 To s)
    ' copy values, empty cells are ignored
    For i = 1 To s
        If (cell_range.Cells(i, 1) <> 0) Then
            t = t + 1
            list(t) = cell_range.Cells(i, 1)
        End If
    Next i
    ReDim Preserve list(1 To t)
End Sub

'----------------------------------------------
'   ExportCellRange
'----------------------------------------------
' build a 2D array to be returned as a cell range
Function ExportCellRange(list As Variant, Optional descending As Boolean = False) As Variant
    Dim i As Long
    Dim resultArray As Variant
    Dim s As Long
    s = UBound(list)
    ReDim resultArray(1 To s, 1 To 1)
    If descending Then
        For i = 1 To s
            resultArray(s + 1 - i, 1) = list(i)
        Next i
    Else
        For i = 1 To s
            resultArray(i, 1) = list(i)
        Next i
    End If
    ExportCellRange = resultArray
End Function

'----------------------------------------------
'   QuickSort
'----------------------------------------------
' a concise and efficient implementation of quick sort, found somewhere on the net
' thanks to the author, I couldn't find his/her name
Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub
