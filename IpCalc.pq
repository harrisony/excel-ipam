let
//==============================================
//   Internal Helper Functions
//==============================================

/*--------------------------------------------
   _IpToNbr
----------------------------------------------
 Converts a list of 4 bytes to a number
 example:
   _IpToNbr({1, 2, 3, 4}) returns 16909060
*/
_IpToNbr = (bytes as list) as number =>
    List.Accumulate(
        bytes,
        0,
        (state, current) => state * 256 + current
    ),

/*--------------------------------------------
   _IpFromNbr
----------------------------------------------
 Converts a number to a list of 4 bytes
 example:
   _IpFromNbr(16909060) returns {1, 2, 3, 4}
*/
_IpFromNbr = (nbr as number) as list =>
    let
        Generate = List.Generate(
            () => [value = nbr, result = {}],
            each List.Count([result]) < 4,
            each [
                value = Number.IntegerDivide([value], 256),
                result = {Number.Mod([value], 256)} & [result]
            ],
            each [result]
        )
    in
        Generate{List.Count(Generate) - 1},

/*--------------------------------------------
   _IpAnd
----------------------------------------------
 Bitwise AND on two IP byte lists
*/
_IpAnd = (ip1 as list, ip2 as list) as list =>
    List.Transform(
        {0..3},
        each Number.BitwiseAnd(ip1{_}, ip2{_})
    ),

/*--------------------------------------------
   _IpOr
----------------------------------------------
 Bitwise OR on two IP byte lists
*/
_IpOr = (ip1 as list, ip2 as list) as list =>
    List.Transform(
        {0..3},
        each Number.BitwiseOr(ip1{_}, ip2{_})
    ),

/*--------------------------------------------
   _IpXor
----------------------------------------------
 Bitwise XOR on two IP byte lists
*/
_IpXor = (ip1 as list, ip2 as list) as list =>
    List.Transform(
        {0..3},
        each Number.BitwiseXor(ip1{_}, ip2{_})
    ),

/*--------------------------------------------
   _IpLenToMask
----------------------------------------------
 Converts a prefix length to a mask byte list
 example:
   _IpLenToMask(24) returns {255, 255, 255, 0}
*/
_IpLenToMask = (len as number) as list =>
    let
        maskNbr = Number.BitwiseShiftLeft(
            Number.Power(2, len) - 1,
            32 - len
        )
    in
        _IpFromNbr(maskNbr),

/*--------------------------------------------
   _ParseIpNet
----------------------------------------------
 Parses a subnet string and returns a record with all components
 Supports both CIDR (/24) and dotted mask (255.255.255.0) notations
 example:
   _ParseIpNet("192.168.1.0/24") returns [Ip={192,168,1,0}, Len=24, Mask={255,255,255,0}, Cidr=true, IpText="192.168.1.0"]
   _ParseIpNet("192.168.1.0 255.255.255.0") returns [Ip={192,168,1,0}, Len=24, Mask={255,255,255,0}, Cidr=false, IpText="192.168.1.0", MaskText="255.255.255.0"]
*/
_ParseIpNet = (s as text) as record =>
    let
        slashPos = Text.PositionOf(s, "/"),
        spacePos = Text.PositionOf(s, " "),

        parsed = if slashPos >= 0 then
            // CIDR notation: "192.168.1.0/24"
            [
                IpText = Text.Start(s, slashPos),
                Len = Number.FromText(Text.Middle(s, slashPos + 1)),
                Cidr = true
            ]
        else if spacePos >= 0 then
            // Dotted mask notation: "192.168.1.0 255.255.255.0"
            let
                ipPart = Text.Start(s, spacePos),
                maskPart = Text.Middle(s, spacePos + 1),
                maskBytes = List.Transform(
                    Text.Split(maskPart, "."),
                    Number.FromText
                ),
                maskNbr = _IpToNbr(maskBytes),
                // Calculate prefix length from mask
                bits = Number.ToText(maskNbr, 2),
                len = Text.Length(bits) - Text.Length(Text.Replace(bits, "1", ""))
            in
                [
                    IpText = ipPart,
                    MaskText = maskPart,
                    Len = len,
                    Cidr = false,
                    MaskBytes = maskBytes
                ]
        else
            // Plain IP address (no mask)
            [
                IpText = s,
                Len = 32,
                Cidr = true
            ],

        // Convert IP text to byte list
        ipBytes = List.Transform(
            Text.Split(parsed[IpText], "."),
            Number.FromText
        ),

        // Get mask
        mask = if Record.HasFields(parsed, "MaskBytes")
               then parsed[MaskBytes]
               else _IpLenToMask(parsed[Len]),

        // Clear host bits
        clearedIp = _IpAnd(ipBytes, mask)
    in
        parsed & [Ip = clearedIp, Mask = mask],

/*--------------------------------------------
   _IpNetToString
----------------------------------------------
 Converts an IpNet record back to string representation
 Preserves the original notation (CIDR vs dotted mask)
*/
_IpNetToString = (net as record) as text =>
    let
        ipText = Text.Combine(
            List.Transform(net[Ip], Number.ToText),
            "."
        ),
        maskPart = if net[Cidr] then
            "/" & Number.ToText(net[Len])
        else
            " " & Text.Combine(
                List.Transform(net[Mask], Number.ToText),
                "."
            )
    in
        ipText & maskPart,

/*--------------------------------------------
   _CompareIpNet
----------------------------------------------
 Compare two IpNet records for sorting
 If IPs are equal, shorter prefix comes first
*/
_CompareIpNet = (a as record, b as record) as number =>
    let
        diff = _IpToNbr(a[Ip]) - _IpToNbr(b[Ip])
    in
        if diff <> 0 then diff else a[Len] - b[Len],

//==============================================
//   IPv4 Public Functions
//==============================================

/*--------------------------------------------
   IpIsValid
----------------------------------------------
 Returns true if an IP address is formatted correctly:
 no space, no extra zeros, no incorrect value
 example:
   IpIsValid("192.168.1.1") returns true
   IpIsValid("192.168.001.1") returns false (leading zero)
   IpIsValid("192.168.1") returns false (missing octet)
*/
IpIsValid = (ip as text) as logical =>
    try
        let
            parts = Text.Split(ip, "."),
            hasCorrectCount = List.Count(parts) = 4,
            roundtrip = IpBinToStr(IpStrToNbr(ip))
        in
            hasCorrectCount and roundtrip = ip
    otherwise
        false,

/*--------------------------------------------
   IpStrToNbr
----------------------------------------------
 Converts a text IP address to number
 example:
   IpStrToNbr("1.2.3.4") returns 16909060
*/
IpStrToNbr = (ip as text) as number =>
    let
        bytes = List.Transform(
            Text.Split(ip, "."),
            Number.FromText
        )
    in
        _IpToNbr(bytes),

/*--------------------------------------------
   IpBinToStr
----------------------------------------------
 Converts a number to text IP address
 example:
   IpBinToStr(16909060) returns "1.2.3.4"
*/
IpBinToStr = (nbr as number) as text =>
    let
        bytes = _IpFromNbr(nbr)
    in
        Text.Combine(
            List.Transform(bytes, Number.ToText),
            "."
        ),

/*--------------------------------------------
   IpAnd
----------------------------------------------
 Bitwise AND operation
 example:
   IpAnd("192.168.1.1", "255.255.255.0") returns "192.168.1.0"
*/
IpAnd = (ip1 as text, ip2 as text) as text =>
    let
        bytes1 = List.Transform(Text.Split(ip1, "."), Number.FromText),
        bytes2 = List.Transform(Text.Split(ip2, "."), Number.FromText),
        result = _IpAnd(bytes1, bytes2)
    in
        Text.Combine(List.Transform(result, Number.ToText), "."),

/*--------------------------------------------
   IpOr
----------------------------------------------
 Bitwise OR operation
 example:
   IpOr("192.168.1.1", "0.0.0.255") returns "192.168.1.255"
*/
IpOr = (ip1 as text, ip2 as text) as text =>
    let
        bytes1 = List.Transform(Text.Split(ip1, "."), Number.FromText),
        bytes2 = List.Transform(Text.Split(ip2, "."), Number.FromText),
        result = _IpOr(bytes1, bytes2)
    in
        Text.Combine(List.Transform(result, Number.ToText), "."),

/*--------------------------------------------
   IpXor
----------------------------------------------
 Bitwise XOR operation
 example:
   IpXor("192.168.1.1", "0.0.0.255") returns "192.168.1.254"
*/
IpXor = (ip1 as text, ip2 as text) as text =>
    let
        bytes1 = List.Transform(Text.Split(ip1, "."), Number.FromText),
        bytes2 = List.Transform(Text.Split(ip2, "."), Number.FromText),
        result = _IpXor(bytes1, bytes2)
    in
        Text.Combine(List.Transform(result, Number.ToText), "."),

/*--------------------------------------------
   IpWithoutMask
----------------------------------------------
 Removes the netmask notation at the end of the IP
 example:
   IpWithoutMask("192.168.1.0/24") returns "192.168.1.0"
   IpWithoutMask("192.168.1.0 255.255.255.0") returns "192.168.1.0"
*/
IpWithoutMask = (s as text) as text =>
    let
        slashPos = Text.PositionOf(s, "/"),
        spacePos = Text.PositionOf(s, " "),
        result = if slashPos >= 0 then
            Text.Start(s, slashPos)
        else if spacePos >= 0 then
            Text.Start(s, spacePos)
        else
            s
    in
        result,

/*--------------------------------------------
   IpSubnetLen
----------------------------------------------
 Get the mask length from a subnet
 example:
   IpSubnetLen("192.168.1.1/24") returns 24
   IpSubnetLen("192.168.1.1 255.255.255.0") returns 24
   IpSubnetLen("192.168.1.1") returns 32
*/
IpSubnetLen = (s as text) as number =>
    let
        net = _ParseIpNet(s)
    in
        net[Len],

/*--------------------------------------------
   IpMask
----------------------------------------------
 Returns an IP netmask from a subnet
 Both notations are accepted
 example:
   IpMask("192.168.1.1/24") returns "255.255.255.0"
   IpMask("192.168.1.1 255.255.255.0") returns "255.255.255.0"
*/
IpMask = (s as text) as text =>
    let
        net = _ParseIpNet(s),
        maskText = Text.Combine(
            List.Transform(net[Mask], Number.ToText),
            "."
        )
    in
        maskText,

/*--------------------------------------------
   IpWildMask
----------------------------------------------
 Returns an IP Wildcard (inverse) mask from a subnet
 Both notations are accepted
 example:
   IpWildMask("192.168.1.1/24") returns "0.0.0.255"
   IpWildMask("192.168.1.1 255.255.255.0") returns "0.0.0.255"
*/
IpWildMask = (s as text) as text =>
    let
        net = _ParseIpNet(s),
        wildNbr = Number.Power(2, 32 - net[Len]) - 1,
        wildBytes = _IpFromNbr(wildNbr)
    in
        Text.Combine(List.Transform(wildBytes, Number.ToText), "."),

/*--------------------------------------------
   IpInvertMask
----------------------------------------------
 Returns an IP Wildcard (inverse) mask from a subnet mask
 or a subnet mask from a wildcard mask
 example:
   IpInvertMask("255.255.255.0") returns "0.0.0.255"
   IpInvertMask("0.0.0.255") returns "255.255.255.0"
*/
IpInvertMask = (mask as text) as text =>
    let
        maskNbr = IpStrToNbr(mask),
        invertedNbr = Number.Power(2, 32) - 1 - maskNbr
    in
        IpBinToStr(invertedNbr),

/*--------------------------------------------
   IpMaskLen
----------------------------------------------
 Returns prefix length from a mask given by string notation (xx.xx.xx.xx)
 example:
   IpMaskLen("255.255.255.0") returns 24
*/
IpMaskLen = (mask as text) as number =>
    let
        maskNbr = IpStrToNbr(mask),
        notMask = Number.Power(2, 32) - 1 - maskNbr,

        // Count trailing zeros
        CountZeros = List.Accumulate(
            {0..31},
            [count = 0, remaining = notMask],
            (state, bit) =>
                if state[remaining] > 0 then
                    [
                        count = state[count] + 1,
                        remaining = Number.IntegerDivide(state[remaining], 2)
                    ]
                else
                    state
        )
    in
        32 - CountZeros[count],

/*--------------------------------------------
   IpAdd
----------------------------------------------
 Add offset to IP address
 example:
   IpAdd("192.168.1.1", 4) returns "192.168.1.5"
   IpAdd("192.168.1.1", 256) returns "192.168.2.1"
*/
IpAdd = (ip as text, offset as number) as text =>
    IpBinToStr(IpStrToNbr(ip) + offset),

/*--------------------------------------------
   IpDiff
----------------------------------------------
 Calculate difference between two IP addresses
 example:
   IpDiff("192.168.1.7", "192.168.1.1") returns 6
*/
IpDiff = (ip1 as text, ip2 as text) as number =>
    IpStrToNbr(ip1) - IpStrToNbr(ip2),

/*--------------------------------------------
   IpGetByte
----------------------------------------------
 Get one byte from an IP address given its position (1-4)
 example:
   IpGetByte("192.168.1.1", 1) returns 192
   IpGetByte("192.168.1.1", 4) returns 1
*/
IpGetByte = (ip as text, pos as number) as number =>
    let
        bytes = List.Transform(Text.Split(ip, "."), Number.FromText)
    in
        bytes{pos - 1},

/*--------------------------------------------
   IpSetByte
----------------------------------------------
 Set one byte in an IP address given its position and value
 example:
   IpSetByte("192.168.1.1", 4, 20) returns "192.168.1.20"
*/
IpSetByte = (ip as text, pos as number, newvalue as number) as text =>
    let
        bytes = List.Transform(Text.Split(ip, "."), Number.FromText),
        updated = List.ReplaceRange(bytes, pos - 1, 1, {newvalue})
    in
        Text.Combine(List.Transform(updated, Number.ToText), "."),

/*--------------------------------------------
   IpSubnetSize
----------------------------------------------
 Returns the number of addresses in a subnet
 example:
   IpSubnetSize("192.168.1.32/29") returns 8
   IpSubnetSize("192.168.1.0 255.255.255.0") returns 256
*/
IpSubnetSize = (subnet as text) as number =>
    let
        net = _ParseIpNet(subnet)
    in
        Number.Power(2, 32 - net[Len]),

/*--------------------------------------------
   IpClearHostBits
----------------------------------------------
 Set to zero the bits in the host part of an address
 example:
   IpClearHostBits("192.168.1.1/24") returns "192.168.1.0/24"
   IpClearHostBits("192.168.1.193 255.255.255.128") returns "192.168.1.128 255.255.255.128"
*/
IpClearHostBits = (net as text) as text =>
    _IpNetToString(_ParseIpNet(net)),

/*--------------------------------------------
   IpIsInSubnet
----------------------------------------------
 Returns TRUE if "ip" is in "subnet"
 example:
   IpIsInSubnet("192.168.1.35", "192.168.1.32/29") returns true
   IpIsInSubnet("192.168.1.35", "192.168.1.32 255.255.255.248") returns true
   IpIsInSubnet("192.168.1.41", "192.168.1.32/29") returns false
*/
IpIsInSubnet = (ip as text, subnet as text) as logical =>
    let
        net = _ParseIpNet(subnet),
        ipBytes = List.Transform(Text.Split(ip, "."), Number.FromText),
        maskedIp = _IpAnd(ipBytes, net[Mask])
    in
        maskedIp = net[Ip],

/*--------------------------------------------
   IpSubnetIsInSubnet
----------------------------------------------
 Returns TRUE if "subnet1" is in "subnet2"
 example:
   IpSubnetIsInSubnet("192.168.1.35/30", "192.168.1.32/29") returns true
   IpSubnetIsInSubnet("192.168.1.41/30", "192.168.1.32/29") returns false
   IpSubnetIsInSubnet("192.168.1.35/28", "192.168.1.32/29") returns false
*/
IpSubnetIsInSubnet = (subnet1 as text, subnet2 as text) as logical =>
    let
        net1 = _ParseIpNet(subnet1),
        net2 = _ParseIpNet(subnet2)
    in
        if net1[Len] < net2[Len] then
            false
        else
            net1[Ip] = _IpAnd(net1[Ip], net2[Mask]),

/*--------------------------------------------
   IpSubnetMatch
----------------------------------------------
 Tries to match an IP address or subnet against a list of subnets and returns the row number
 Returns the smallest matching subnet (best match)
 example:
   IpSubnetMatch("192.168.1.35", {{"192.168.1.0/24"}, {"192.168.1.32/29"}}) returns 2
*/
IpSubnetMatch = (ip as text, table as list) as number =>
    let
        searchNet = _ParseIpNet(ip),

        // Find best match (smallest subnet that contains the IP)
        FindBest = List.Accumulate(
            List.Positions(table),
            [bestRow = 0, bestLen = -1],
            (state, idx) =>
                let
                    subnetText = if List.Count(table{idx}) > 0 then
                        table{idx}{0}
                    else
                        table{idx},
                    net = _ParseIpNet(subnetText)
                in
                    if net[Len] > state[bestLen] and
                       searchNet[Len] >= net[Len] and
                       _IpAnd(searchNet[Ip], net[Mask]) = net[Ip] then
                        [bestRow = idx + 1, bestLen = net[Len]]
                    else
                        state
        )
    in
        FindBest[bestRow],

/*--------------------------------------------
   IpSubnetVLookup
----------------------------------------------
 Tries to match an IP address or subnet against a list of subnets
 and returns the value in the same row based on the index_number
 example:
   IpSubnetVLookup("192.168.1.35", {{"192.168.1.0/24", "LAN"}, {"10.0.0.0/8", "Private"}}, 2) returns "LAN"
*/
IpSubnetVLookup = (ip as text, table as list, index as number) as any =>
    let
        row = IpSubnetMatch(ip, table)
    in
        if row = 0 then
            "Not Found"
        else
            table{row - 1}{index - 1},

/*--------------------------------------------
   IpIsPrivate
----------------------------------------------
 Returns TRUE if "ip" is in one of the private IP address ranges
 example:
   IpIsPrivate("192.168.1.35") returns true
   IpIsPrivate("209.85.148.104") returns false
*/
IpIsPrivate = (ip as text) as logical =>
    IpIsInSubnet(ip, "10.0.0.0/8") or
    IpIsInSubnet(ip, "172.16.0.0/12") or
    IpIsInSubnet(ip, "192.168.0.0/16"),

/*--------------------------------------------
   IpDivideSubnet
----------------------------------------------
 Divide a network into smaller subnets
 "n" is the value that will be added to the subnet length
 Returns a list of smaller subnets
 example:
   IpDivideSubnet("1.2.3.0/24", 2) returns {"1.2.3.0/26", "1.2.3.64/26", "1.2.3.128/26", "1.2.3.192/26"}
*/
IpDivideSubnet = (subnet as text, n as number) as list =>
    let
        net = _ParseIpNet(subnet),
        newLen = net[Len] + n,

        result = if newLen > 32 then
            error "ERR subnet length > 32"
        else
            let
                subnetSize = Number.Power(2, 32 - newLen),
                count = Number.Power(2, n),
                startNbr = _IpToNbr(net[Ip]),

                subnets = List.Transform(
                    {0..count - 1},
                    each
                        let
                            ipNbr = startNbr + (_ * subnetSize),
                            ipText = IpBinToStr(ipNbr),
                            subnetText = if net[Cidr] then
                                ipText & "/" & Number.ToText(newLen)
                            else
                                ipText & " " & IpMask(ipText & "/" & Number.ToText(newLen))
                        in
                            subnetText
                )
            in
                subnets
    in
        result,

/*--------------------------------------------
   IpRangeToCIDR
----------------------------------------------
 Returns a network or a list of networks given the first and last address of an IP range
 example:
   IpRangeToCIDR("10.0.0.1", "10.0.0.254") returns {"10.0.0.0/24"}
   IpRangeToCIDR("10.0.0.1", "10.0.1.63") returns {"10.0.0.0/24", "10.0.1.0/26"}
*/
IpRangeToCIDR = (firstAddr as text, lastAddr as text) as list =>
    let
        // Set last bit to zero for first address, to one for last address
        first = IpAnd(firstAddr, "255.255.255.254"),
        last = IpOr(lastAddr, "0.0.0.1"),

        // Recursive function to build network list
        BuildNetworks = (currentFirst as text, targetLast as text) as list =>
            let
                // Find the largest network starting at currentFirst that doesn't exceed targetLast
                FindNetwork = List.Accumulate(
                    {1..32},
                    [len = 32, ip1 = currentFirst, ip2 = currentFirst, diff = IpDiff(targetLast, currentFirst)],
                    (state, testLen) =>
                        let
                            net = _ParseIpNet(currentFirst & "/" & Number.ToText(testLen)),
                            netFirst = Text.Combine(List.Transform(net[Ip], Number.ToText), "."),
                            wildBytes = _IpFromNbr(Number.Power(2, 32 - testLen) - 1),
                            netLast = Text.Combine(List.Transform(_IpOr(net[Ip], wildBytes), Number.ToText), "."),
                            newDiff = IpDiff(netLast, targetLast)
                        in
                            if netFirst = currentFirst and newDiff <= 0 then
                                [len = testLen, ip1 = netFirst, ip2 = netLast, diff = newDiff]
                            else
                                state
                ),

                netText = FindNetwork[ip1] & "/" & Number.ToText(FindNetwork[len])
            in
                if FindNetwork[diff] = 0 then
                    {netText}
                else
                    {netText} & @BuildNetworks(IpAdd(FindNetwork[ip2], 1), targetLast)
    in
        BuildNetworks(first, last),

//==============================================
//   Array Functions
//==============================================

/*--------------------------------------------
   IpSortArray
----------------------------------------------
 Sort an array of IP addresses
 example:
   IpSortArray({"192.168.1.10", "10.0.0.1", "192.168.1.5"}) returns {"10.0.0.1", "192.168.1.5", "192.168.1.10"}
*/
IpSortArray = (arr as list) as list =>
    let
        // Remove empty values
        filtered = List.Select(arr, each _ <> null and _ <> ""),

        // Sort by converting to numbers
        sorted = List.Sort(
            filtered,
            (a, b) =>
                let
                    aNbr = IpStrToNbr(a),
                    bNbr = IpStrToNbr(b)
                in
                    if aNbr < bNbr then -1 else if aNbr > bNbr then 1 else 0
        )
    in
        sorted,

/*--------------------------------------------
   IpSubnetSortArray
----------------------------------------------
 Sort an array of IP subnets
 example:
   IpSubnetSortArray({"192.168.1.0/24", "10.0.0.0/8", "192.168.1.32/29"})
*/
IpSubnetSortArray = (arr as list) as list =>
    let
        // Remove empty values and parse
        filtered = List.Select(arr, each _ <> null and _ <> ""),
        nets = List.Transform(filtered, _ParseIpNet),

        // Sort
        sorted = List.Sort(nets, _CompareIpNet),

        // Convert back to strings
        result = List.Transform(sorted, _IpNetToString)
    in
        result,

/*--------------------------------------------
   IpFindOverlappingSubnets
----------------------------------------------
 Find overlapping subnets in a list
 Returns a list where each element shows the larger subnet that contains it, or empty if no overlap
 example:
   IpFindOverlappingSubnets({"192.168.1.0/24", "192.168.1.32/29", "10.0.0.0/8"})
   returns {"", "192.168.1.0/24", ""}
*/
IpFindOverlappingSubnets = (arr as list) as list =>
    let
        filtered = List.Select(arr, each _ <> null and _ <> ""),
        nets = List.Transform(filtered, _ParseIpNet),

        // For each subnet, find if it's contained in another
        results = List.Transform(
            {0..List.Count(nets) - 1},
            (i) =>
                let
                    netI = nets{i},
                    // Find first containing subnet
                    containing = List.First(
                        List.Select(
                            {0..List.Count(nets) - 1},
                            (j) =>
                                let
                                    netJ = nets{j}
                                in
                                    i <> j and
                                    netI[Len] > netJ[Len] and
                                    _IpAnd(netI[Ip], netJ[Mask]) = netJ[Ip]
                        ),
                        null
                    )
                in
                    if containing = null then "" else _IpNetToString(nets{containing})
        )
    in
        results,

/*--------------------------------------------
   IpSubnetAggregateArray
----------------------------------------------
 Sort and aggregate subnets:
 - Remove subnets included in larger subnets
 - Join contiguous subnets
 - Remove duplicates
 example:
   IpSubnetAggregateArray({"192.168.1.0/25", "192.168.1.128/25", "192.168.1.32/29"})
   returns {"192.168.1.0/24"}
*/
IpSubnetAggregateArray = (arr as list) as list =>
    let
        // Remove empty values and parse
        filtered = List.Select(arr, each _ <> null and _ <> ""),
        nets = List.Transform(filtered, _ParseIpNet),

        // Sort
        sorted = List.Sort(nets, _CompareIpNet),

        // Aggregate using recursion
        Aggregate = (netList as list, index as number) as list =>
            if index >= List.Count(netList) - 1 then
                netList
            else
                let
                    current = netList{index},
                    next = netList{index + 1},

                    // Check if next is included in current
                    nextIncluded = next[Len] >= current[Len] and
                                   _IpAnd(next[Ip], current[Mask]) = current[Ip],

                    // Check if can join (same length and contiguous)
                    canJoin = if not nextIncluded and current[Len] = next[Len] and current[Len] > 0 then
                        let
                            // Try creating a subnet with mask 1 bit shorter
                            bignetLen = current[Len] - 1,
                            bignetMask = _IpLenToMask(bignetLen),
                            bignetIp = _IpAnd(current[Ip], bignetMask)
                        in
                            _IpAnd(next[Ip], bignetMask) = bignetIp
                    else
                        false,

                    newList = if nextIncluded then
                        // Remove next
                        List.RemoveRange(netList, index + 1, 1)
                    else if canJoin then
                        // Join current and next into a larger subnet
                        let
                            bignetLen = current[Len] - 1,
                            bignetMask = _IpLenToMask(bignetLen),
                            bignetIp = _IpAnd(current[Ip], bignetMask),
                            bignet = [
                                Ip = bignetIp,
                                Len = bignetLen,
                                Mask = bignetMask,
                                Cidr = current[Cidr]
                            ]
                        in
                            List.ReplaceRange(
                                List.RemoveRange(netList, index + 1, 1),
                                index,
                                1,
                                {bignet}
                            )
                    else
                        netList,

                    nextIndex = if nextIncluded or canJoin then
                        // Stay at current index to check if we can aggregate more
                        if index > 0 then index - 1 else index
                    else
                        index + 1
                in
                    @Aggregate(newList, nextIndex),

        aggregated = Aggregate(sorted, 0),

        // Convert back to strings
        result = List.Transform(aggregated, _IpNetToString)
    in
        result,



/*--------------------------------------------
   IpSubtractSubnets
----------------------------------------------
 Remove subnets from a list of subnets
 Returns the unused portions
 example:
   IpSubtractSubnets({"192.168.1.0/24"}, {"192.168.1.0/25"}) returns {"192.168.1.128/25"}
*/
IpSubtractSubnets = (inputArray as list, subtractArray as list) as list =>
    let
        // Aggregate both arrays first
        input = IpSubnetAggregateArray(inputArray),
        subtract = IpSubnetAggregateArray(subtractArray),

        inputNets = List.Transform(input, _ParseIpNet),
        subtractNets = List.Transform(subtract, _ParseIpNet),

        // Process
        Subtract = (inList as list, subList as list, i as number, j as number) as list =>
            if i >= List.Count(subList) or j >= List.Count(inList) then
                inList
            else
                let
                    subNet = subList{i},
                    inNet = inList{j},

                    // Check if subtract net contains input net
                    subContainsIn = inNet[Len] >= subNet[Len] and
                                    _IpAnd(inNet[Ip], subNet[Mask]) = subNet[Ip],

                    // Check if input net contains subtract net
                    inContainsSub = subNet[Len] > inNet[Len] and
                                    _IpAnd(subNet[Ip], inNet[Mask]) = inNet[Ip],

                    newState = if subContainsIn then
                        // Remove input net
                        [newList = List.RemoveRange(inList, j, 1), newI = i, newJ = j]
                    else if inContainsSub then
                        // Split input net
                        let
                            split = IpDivideSubnet(_IpNetToString(inNet), 1),
                            splitNets = List.Transform(split, _ParseIpNet),
                            updated = List.ReplaceRange(inList, j, 1, splitNets)
                        in
                            [newList = updated, newI = i, newJ = j]
                    else
                        // No overlap - advance
                        let
                            inNbr = _IpToNbr(inNet[Ip]),
                            subNbr = _IpToNbr(subNet[Ip])
                        in
                            if inNbr < subNbr then
                                [newList = inList, newI = i, newJ = j + 1]
                            else
                                [newList = inList, newI = i + 1, newJ = j]
                in
                    @Subtract(newState[newList], subList, newState[newI], newState[newJ]),

        result = Subtract(inputNets, subtractNets, 0, 0),
        resultText = List.Transform(result, _IpNetToString)
    in
        resultText,

/*--------------------------------------------
   IpCommonSubnets
----------------------------------------------
 Returns the list of IP ranges which are common to both input lists
 example:
   IpCommonSubnets({"192.168.0.0/23"}, {"192.168.1.0/24"}) returns {"192.168.1.0/24"}
*/
IpCommonSubnets = (inputArray1 as list, inputArray2 as list) as list =>
    let
        diff = IpSubtractSubnets(inputArray1, inputArray2)
    in
        IpSubtractSubnets(inputArray1, diff),

/*--------------------------------------------
   IpMergeOnSubnet
----------------------------------------------
 Merges two tables based on subnet containment
 Joins rows where an IP address in leftTable falls within a subnet range in rightTable
 Returns the most specific (longest prefix) match when multiple subnets contain an IP
 example:
   IpMergeOnSubnet(ipTable, "IP", subnetTable, "Subnet", "Left")
   IpMergeOnSubnet(ipTable, "IP", subnetTable, "Subnet", "Inner")
*/
IpMergeOnSubnet = (
    leftTable as table,
    leftIpColumn as text,
    rightTable as table,
    rightSubnetColumn as text,
    optional joinKind as text
) as table =>
    let
        // Default to Left join if not specified
        actualJoinKind = if joinKind = null then "Left" else joinKind,

        // Validate join kind
        validJoinKinds = {"Left", "Inner", "Right", "Full"},
        isValidJoinKind = List.Contains(validJoinKinds, actualJoinKind),

        result = if not isValidJoinKind then
            error "Invalid joinKind. Must be one of: Left, Inner, Right, Full"
        else
            let
                // Add index columns for tracking
                leftWithIndex = Table.AddIndexColumn(leftTable, "__LeftIndex", 0, 1),
                rightWithIndex = Table.AddIndexColumn(rightTable, "__RightIndex", 0, 1),

                // Get column lists for later merging
                leftColumns = Table.ColumnNames(leftTable),
                rightColumns = Table.ColumnNames(rightTable),

                // For each row in left table, find the best matching subnet in right table
                leftWithMatch = Table.AddColumn(
                    leftWithIndex,
                    "__MatchIndex",
                    (leftRow) =>
                        let
                            ipValue = Record.Field(leftRow, leftIpColumn),

                            // Handle null or empty IP
                            matchIndex = if ipValue = null or ipValue = "" then
                                null
                            else
                                try
                                    let
                                        ipBytes = List.Transform(
                                            Text.Split(ipValue, "."),
                                            Number.FromText
                                        ),

                                        // Find best match (most specific subnet)
                                        FindBest = List.Accumulate(
                                            Table.ToRecords(rightWithIndex),
                                            [bestIndex = null, bestLen = -1],
                                            (state, rightRow) =>
                                                let
                                                    subnetValue = Record.Field(rightRow, rightSubnetColumn)
                                                in
                                                    if subnetValue = null or subnetValue = "" then
                                                        state
                                                    else
                                                        try
                                                            let
                                                                net = _ParseIpNet(subnetValue),
                                                                maskedIp = _IpAnd(ipBytes, net[Mask]),
                                                                isMatch = maskedIp = net[Ip]
                                                            in
                                                                if isMatch and net[Len] > state[bestLen] then
                                                                    [
                                                                        bestIndex = Record.Field(rightRow, "__RightIndex"),
                                                                        bestLen = net[Len]
                                                                    ]
                                                                else
                                                                    state
                                                        otherwise
                                                            state
                                        )
                                    in
                                        FindBest[bestIndex]
                                otherwise
                                    null
                        in
                            matchIndex,
                    type nullable number
                ),

                // Build result based on join kind
                finalResult = if actualJoinKind = "Inner" then
                    // Inner join: only rows with matches
                    let
                        matched = Table.SelectRows(leftWithMatch, each [__MatchIndex] <> null),
                        joined = Table.NestedJoin(
                            matched,
                            {"__MatchIndex"},
                            rightWithIndex,
                            {"__RightIndex"},
                            "__RightData",
                            JoinKind.Inner
                        ),
                        expanded = Table.ExpandTableColumn(
                            joined,
                            "__RightData",
                            rightColumns,
                            rightColumns
                        ),
                        cleaned = Table.RemoveColumns(expanded, {"__LeftIndex", "__RightIndex", "__MatchIndex"})
                    in
                        cleaned

                else if actualJoinKind = "Left" then
                    // Left join: all left rows, nulls for non-matches
                    let
                        joined = Table.NestedJoin(
                            leftWithMatch,
                            {"__MatchIndex"},
                            rightWithIndex,
                            {"__RightIndex"},
                            "__RightData",
                            JoinKind.LeftOuter
                        ),
                        expanded = Table.ExpandTableColumn(
                            joined,
                            "__RightData",
                            rightColumns,
                            rightColumns
                        ),
                        cleaned = Table.RemoveColumns(expanded, {"__LeftIndex", "__RightIndex", "__MatchIndex"})
                    in
                        cleaned

                else if actualJoinKind = "Right" then
                    // Right join: all right rows, nulls for non-matches
                    let
                        // Find which right rows have matches
                        matchedRightIndexes = List.Distinct(
                            List.Select(
                                Table.Column(leftWithMatch, "__MatchIndex"),
                                each _ <> null
                            )
                        ),

                        rightWithMatchFlag = Table.AddColumn(
                            rightWithIndex,
                            "__HasMatch",
                            each List.Contains(matchedRightIndexes, [__RightIndex])
                        ),

                        // For matched right rows, get corresponding left rows
                        matchedPairs = Table.SelectRows(
                            leftWithMatch,
                            each [__MatchIndex] <> null
                        ),

                        joined = Table.NestedJoin(
                            rightWithMatchFlag,
                            {"__RightIndex"},
                            matchedPairs,
                            {"__MatchIndex"},
                            "__LeftData",
                            JoinKind.LeftOuter
                        ),

                        // Expand left columns
                        expanded = Table.ExpandTableColumn(
                            joined,
                            "__LeftData",
                            leftColumns,
                            leftColumns
                        ),

                        cleaned = Table.RemoveColumns(expanded, {"__RightIndex", "__HasMatch"})
                    in
                        cleaned

                else // Full join
                    // Full join: all rows from both tables
                    let
                        // Get Left join result
                        leftJoin = Table.NestedJoin(
                            leftWithMatch,
                            {"__MatchIndex"},
                            rightWithIndex,
                            {"__RightIndex"},
                            "__RightData",
                            JoinKind.LeftOuter
                        ),
                        leftExpanded = Table.ExpandTableColumn(
                            leftJoin,
                            "__RightData",
                            rightColumns,
                            rightColumns
                        ),

                        // Find unmatched right rows
                        matchedRightIndexes = List.Distinct(
                            List.Select(
                                Table.Column(leftWithMatch, "__MatchIndex"),
                                each _ <> null
                            )
                        ),

                        unmatchedRight = Table.SelectRows(
                            rightWithIndex,
                            each not List.Contains(matchedRightIndexes, [__RightIndex])
                        ),

                        // Create null columns for left side
                        unmatchedRightWithNulls = List.Accumulate(
                            leftColumns,
                            unmatchedRight,
                            (tbl, colName) => Table.AddColumn(tbl, colName, each null)
                        ),

                        // Combine both
                        combined = Table.Combine({
                            leftExpanded,
                            unmatchedRightWithNulls
                        }),

                        cleaned = Table.RemoveColumns(combined, {"__LeftIndex", "__RightIndex", "__MatchIndex"})
                    in
                        cleaned
            in
                finalResult
    in
        result,

//==============================================
//   IPv6 Internal Helper Functions
//==============================================

/*--------------------------------------------
   _Ipv6Parse
----------------------------------------------
 Parse an IPv6 address into a record with blocks array and length
*/
_Ipv6Parse = (s as text) as record =>
    let
        parts = Text.Split(s, "/"),
        len = if List.Count(parts) > 1 then
            Number.FromText(parts{1})
        else
            128,

        ipText = parts{0},
        blocks = Text.Split(ipText, ":"),

        // Find position of empty block (::)
        emptyPos = List.PositionOf(blocks, ""),

        // Expand :: notation
        expanded = if emptyPos >= 0 then
            let
                before = List.FirstN(blocks, emptyPos),
                after = List.Skip(blocks, emptyPos + 1),
                // Remove any additional empty strings
                afterClean = List.Select(after, each _ <> ""),
                zerosCount = 8 - List.Count(before) - List.Count(afterClean),
                zeros = List.Repeat({"0"}, zerosCount)
            in
                before & zeros & afterClean
        else
            blocks,

        // Convert to numbers
        ar = List.Transform(
            expanded,
            each
                try Number.FromText("0x" & _)
                otherwise 0
        ),

        // Ensure exactly 8 blocks
        final = if List.Count(ar) < 8 then
            ar & List.Repeat({0}, 8 - List.Count(ar))
        else if List.Count(ar) > 8 then
            List.FirstN(ar, 8)
        else
            ar
    in
        [Ar = final, Len = len],

/*--------------------------------------------
   _Ipv6Long
----------------------------------------------
 Get the long form (expanded) IPv6 address
*/
_Ipv6Long = (rec as record) as text =>
    let
        parts = List.Transform(
            rec[Ar],
            each Text.PadStart(Number.ToText(_, "X"), 4, "0")
        )
    in
        Text.Lower(Text.Combine(parts, ":")),

/*--------------------------------------------
   _Ipv6Compress
----------------------------------------------
 Get the compressed form of IPv6 address
*/
_Ipv6Compress = (rec as record) as text =>
    let
        // Convert to hex strings
        parts = List.Transform(rec[Ar], each Number.ToText(_, "X")),
        joined = Text.Combine(parts, ":"),

        // Try to replace longest sequence of zeros with ::
        compressed = List.Accumulate(
            {":0:0:0:0:0:0:", ":0:0:0:0:0:", ":0:0:0:0:", ":0:0:0:", ":0:0:"},
            joined,
            (state, pattern) =>
                if Text.Contains(state, pattern) then
                    Text.Replace(state, pattern, "::", 1)
                else
                    state
        ),

        // Clean up edge cases
        final = Text.Replace(
            Text.Replace(compressed, "0::", "::"),
            "::0",
            "::"
        )
    in
        Text.Lower(final),

/*--------------------------------------------
   _Ipv6ToBin
----------------------------------------------
 Convert IPv6 to 128-character binary string
*/
_Ipv6ToBin = (rec as record) as text =>
    let
        binary = List.Transform(
            rec[Ar],
            each Text.PadStart(Number.ToText(_, 2), 16, "0")
        )
    in
        Text.Combine(binary, ""),

/*--------------------------------------------
   _Ipv6FromBin
----------------------------------------------
 Create IPv6 from 128-character binary string
*/
_Ipv6FromBin = (bin as text) as record =>
    let
        ar = if Text.Length(bin) = 128 then
            List.Transform(
                {0..7},
                each Number.FromText(Text.Middle(bin, _ * 16, 16), 2)
            )
        else
            List.Repeat({0}, 8)
    in
        [Ar = ar, Len = 128],

//==============================================
//   IPv6 Public Functions
//==============================================

/*--------------------------------------------
   Ipv6MaskLen
----------------------------------------------
 Returns prefix length from an IPv6 net
 example:
   Ipv6MaskLen("2001:db8:1f89::/48") returns 48
*/
Ipv6MaskLen = (s as text) as number =>
    let
        slashPos = Text.PositionOf(s, "/")
    in
        if slashPos >= 0 then
            Number.FromText(Text.Middle(s, slashPos + 1))
        else
            128,

/*--------------------------------------------
   Ipv6WithoutMask
----------------------------------------------
 Removes the /xx netmask notation at the end of the IP
 example:
   Ipv6WithoutMask("2001:db8:1f89::/48") returns "2001:db8:1f89::"
*/
Ipv6WithoutMask = (s as text) as text =>
    let
        slashPos = Text.PositionOf(s, "/")
    in
        if slashPos >= 0 then
            Text.Start(s, slashPos)
        else
            s,

/*--------------------------------------------
   Ipv6IsValid
----------------------------------------------
 Returns true if an IPv6 address has a valid format
 example:
   Ipv6IsValid("2001:db8::1") returns true
   Ipv6IsValid("invalid") returns false
*/
Ipv6IsValid = (ip as text) as logical =>
    try
        let
            doubleColonCount = Text.Length(ip) - Text.Length(Text.Replace(ip, "::", "")),
            colonCount = Text.Length(ip) - Text.Length(Text.Replace(ip, ":", "")),

            // Check for invalid characters
            validChars = "0123456789abcdefABCDEF:",
            hasInvalidChars = List.Count(
                List.Select(
                    Text.ToList(ip),
                    each not Text.Contains(validChars, _)
                )
            ) > 0,

            // Parse and check
            parsed = _Ipv6Parse(ip),
            valid = not hasInvalidChars and
                    ((doubleColonCount = 0 and colonCount = 7) or
                     (doubleColonCount = 2 and colonCount <= 7))
        in
            valid
    otherwise
        false,

/*--------------------------------------------
   Ipv6Expand
----------------------------------------------
 Returns a representation of an IPv6 address with all the missing zeros
 The result has a fixed length of 39 characters
 example:
   Ipv6Expand("1:2:3::8") returns "0001:0002:0003:0000:0000:0000:0000:0008"
*/
Ipv6Expand = (ip as text) as text =>
    let
        parsed = _Ipv6Parse(Ipv6WithoutMask(ip))
    in
        _Ipv6Long(parsed),

/*--------------------------------------------
   Ipv6Compress
----------------------------------------------
 Returns the shortest representation of an IPv6 address
 example:
   Ipv6Compress("0001:0002:0003:0000:0000:0000:0000:0008") returns "1:2:3::8"
   Ipv6Compress("01:0:0::") returns "1::"
*/
Ipv6Compress = (ip as text) as text =>
    let
        parsed = _Ipv6Parse(Ipv6WithoutMask(ip))
    in
        _Ipv6Compress(parsed),

/*--------------------------------------------
   Ipv6ToBin
----------------------------------------------
 Returns a string representing the binary value of IPv6 address
 The result has a fixed length of 128 characters
 example:
   Ipv6ToBin("2001:db8::1")
*/
Ipv6ToBin = (ip as text) as text =>
    let
        parsed = _Ipv6Parse(Ipv6WithoutMask(ip))
    in
        _Ipv6ToBin(parsed),

/*--------------------------------------------
   Ipv6FromBin
----------------------------------------------
 Returns an IPv6 from a string representing the binary value of IPv6 address
 The parameter must be a 128 character string
 example:
   Ipv6FromBin("00100000...") returns compressed IPv6
*/
Ipv6FromBin = (bin as text) as text =>
    let
        parsed = _Ipv6FromBin(bin)
    in
        _Ipv6Compress(parsed),

/*--------------------------------------------
   Ipv6AddInt
----------------------------------------------
 Add a value to an IPv6 address
 example:
   Ipv6AddInt("1::2", 16) returns "1::12"
*/
Ipv6AddInt = (ip as text, offset as number) as text =>
    let
        parsed = _Ipv6Parse(Ipv6WithoutMask(ip)),

        // Add offset starting from the rightmost block
        added = List.Accumulate(
            List.Reverse({0..7}),
            [ar = parsed[Ar], carry = offset],
            (state, idx) =>
                if state[carry] = 0 then
                    state
                else
                    let
                        newVal = state[ar]{idx} + state[carry],
                        blockVal = Number.Mod(newVal, 65536),
                        newCarry = Number.IntegerDivide(newVal, 65536),
                        newAr = List.ReplaceRange(state[ar], idx, 1, {blockVal})
                    in
                        [ar = newAr, carry = newCarry]
        )
    in
        _Ipv6Compress([Ar = added[ar], Len = parsed[Len]]),

/*--------------------------------------------
   Ipv6Add
----------------------------------------------
 Add two IPv6 addresses
 example:
   Ipv6Add("1:2::", "::3") returns "1:2::3"
   Ipv6Add("1:2::2", "::3") returns "1:2::5"
*/
Ipv6Add = (ip1 as text, ip2 as text) as text =>
    let
        parsed1 = _Ipv6Parse(Ipv6WithoutMask(ip1)),
        parsed2 = _Ipv6Parse(Ipv6WithoutMask(ip2)),

        // Add blocks with carry
        added = List.Accumulate(
            List.Reverse({0..7}),
            [ar = List.Repeat({0}, 8), carry = 0],
            (state, idx) =>
                let
                    sum = parsed1[Ar]{idx} + parsed2[Ar]{idx} + state[carry],
                    blockVal = Number.Mod(sum, 65536),
                    newCarry = Number.IntegerDivide(sum, 65536),
                    newAr = List.ReplaceRange(state[ar], idx, 1, {blockVal})
                in
                    [ar = newAr, carry = newCarry]
        )
    in
        _Ipv6Compress([Ar = added[ar], Len = 128]),

/*--------------------------------------------
   Ipv6GetBlock
----------------------------------------------
 Returns the 4-digit hexa block at position blockNbr (1-8)
 Block 1 is the block on the left
 example:
   Ipv6GetBlock("2001:db8:1f89:c5a3::ac1f:8001", 2) returns "0db8"
*/
Ipv6GetBlock = (ip as text, blockNbr as number) as text =>
    let
        parsed = _Ipv6Parse(Ipv6WithoutMask(ip)),
        block = parsed[Ar]{blockNbr - 1}
    in
        Text.Lower(Text.PadStart(Number.ToText(block, "X"), 4, "0")),

/*--------------------------------------------
   Ipv6GetBlockInt
----------------------------------------------
 Same as Ipv6GetBlock but returns an integer between 0 and 65535
 example:
   Ipv6GetBlockInt("2001:db8:1f89:c5a3::ac1f:8001", 2) returns 3512
*/
Ipv6GetBlockInt = (ip as text, blockNbr as number) as number =>
    let
        parsed = _Ipv6Parse(Ipv6WithoutMask(ip))
    in
        parsed[Ar]{blockNbr - 1},

/*--------------------------------------------
   Ipv6SetBlock
----------------------------------------------
 Sets the value of the 4-digit hexa block at position blockNbr (1-8)
 example:
   Ipv6SetBlock("2001::", 2, "db8") returns "2001:db8::"
*/
Ipv6SetBlock = (ip as text, blockNbr as number, valHex as text) as text =>
    let
        parsed = _Ipv6Parse(Ipv6WithoutMask(ip)),
        val = try Number.FromText("0x" & valHex) otherwise 0,
        updated = List.ReplaceRange(parsed[Ar], blockNbr - 1, 1, {Number.Mod(val, 65536)})
    in
        _Ipv6Compress([Ar = updated, Len = parsed[Len]]),

/*--------------------------------------------
   Ipv6SetBlockInt
----------------------------------------------
 Same as Ipv6SetBlock but the block value is passed as an integer
 example:
   Ipv6SetBlockInt("2001::", 2, 3512) returns "2001:db8::"
*/
Ipv6SetBlockInt = (ip as text, blockNbr as number, valInt as number) as text =>
    let
        parsed = _Ipv6Parse(Ipv6WithoutMask(ip)),
        updated = List.ReplaceRange(parsed[Ar], blockNbr - 1, 1, {Number.Mod(valInt, 65536)})
    in
        _Ipv6Compress([Ar = updated, Len = parsed[Len]]),

/*--------------------------------------------
   Ipv6SetBits
----------------------------------------------
 Sets one or more bits in an IPv6 address
 bits is a string with one or more "0" and "1"
 offset is the position of the first bit to set (1 to 128 from left to right)
 example:
   Ipv6SetBits("2001:db8::", 65, "1111111111111111") sets bits 65-80
*/
Ipv6SetBits = (ip as text, offset as number, bits as text) as text =>
    let
        parsed = _Ipv6Parse(Ipv6WithoutMask(ip)),
        bin = _Ipv6ToBin(parsed),
        newBin = Text.Start(bin, offset - 1) &
                 bits &
                 (if offset + Text.Length(bits) <= 128 then
                     Text.Middle(bin, offset + Text.Length(bits) - 1)
                  else
                     ""),
        finalBin = Text.Start(newBin, 128)
    in
        Ipv6FromBin(finalBin),

/*--------------------------------------------
   Ipv6GetIpv4
----------------------------------------------
 Get the value of an IPv4 embedded in an IPv6 at a given block position
 example:
   Ipv6GetIpv4("2001:c0a8:102::", 2) returns "192.168.1.2"
*/
Ipv6GetIpv4 = (ipv6 as text, blockNbr as number) as text =>
    let
        parsed = _Ipv6Parse(Ipv6WithoutMask(ipv6)),
        block1 = parsed[Ar]{blockNbr - 1},
        block2 = parsed[Ar]{blockNbr},
        ipv4Nbr = block1 * 65536 + block2
    in
        IpBinToStr(ipv4Nbr),

/*--------------------------------------------
   Ipv6SetIpv4
----------------------------------------------
 Put the value of an IPv4 in an IPv6 at a given block position
 example:
   Ipv6SetIpv4("2001::", 2, "192.168.1.2") returns "2001:c0a8:102::"
*/
Ipv6SetIpv4 = (ipv6 as text, blockNbr as number, ipv4 as text) as text =>
    let
        parsed = _Ipv6Parse(Ipv6WithoutMask(ipv6)),
        ipv4Nbr = IpStrToNbr(ipv4),
        block1 = Number.IntegerDivide(ipv4Nbr, 65536),
        block2 = Number.Mod(ipv4Nbr, 65536),
        updated = List.ReplaceRange(
            List.ReplaceRange(parsed[Ar], blockNbr - 1, 1, {block1}),
            blockNbr,
            1,
            {block2}
        )
    in
        _Ipv6Compress([Ar = updated, Len = parsed[Len]]),

/*--------------------------------------------
   Ipv6IsInSubnet
----------------------------------------------
 Returns TRUE if "ip" is in "subnet"
 example:
   Ipv6IsInSubnet("2001:db8:1::ac1f:1", "2001:db8:1::/48") returns true
   Ipv6IsInSubnet("2001:db8:2::ac1f:1", "2001:db8:1::/48") returns false
*/
Ipv6IsInSubnet = (ip as text, subnet as text) as logical =>
    let
        prefixLen = Ipv6MaskLen(subnet),
        subnetBin = Text.Start(Ipv6ToBin(subnet), prefixLen),
        ipBin = Text.Start(Ipv6ToBin(ip), prefixLen)
    in
        subnetBin = ipBin,

/*--------------------------------------------
   Ipv6SubnetFirstAddress
----------------------------------------------
 Returns the first address of an IPv6 subnet
 example:
   Ipv6SubnetFirstAddress("2001:db8:1:1a0::/59") returns "2001:db8:1:1a0::"
*/
Ipv6SubnetFirstAddress = (subnet as text) as text =>
    let
        prefixLen = Ipv6MaskLen(subnet),
        zeros = Text.Repeat("0", 128 - prefixLen)
    in
        Ipv6SetBits(subnet, prefixLen + 1, zeros),

/*--------------------------------------------
   Ipv6SubnetLastAddress
----------------------------------------------
 Returns the last address of an IPv6 subnet
 example:
   Ipv6SubnetLastAddress("2001:db8:1:1a0::/59") returns "2001:db8:1:1bf:ffff:ffff:ffff:ffff"
*/
Ipv6SubnetLastAddress = (subnet as text) as text =>
    let
        prefixLen = Ipv6MaskLen(subnet),
        ones = Text.Repeat("1", 128 - prefixLen)
    in
        Ipv6SetBits(subnet, prefixLen + 1, ones),

/*--------------------------------------------
   Ipv6SortArray
----------------------------------------------
 Sort an array of IPv6 addresses
 example:
   Ipv6SortArray({"2001:db8::1", "2001:db8::10", "2001:db8::2"})
*/
Ipv6SortArray = (arr as list) as list =>
    let
        // Remove empty values
        filtered = List.Select(arr, each _ <> null and _ <> ""),

        // Sort by expanded form
        sorted = List.Sort(
            filtered,
            (a, b) =>
                let
                    aExp = Ipv6Expand(a),
                    bExp = Ipv6Expand(b)
                in
                    if aExp = bExp then 0
                    else if aExp < bExp then -1
                    else 1
        )
    in
        sorted

/*--------------------------------------------
   IpSortTable
 ----------------------------------------------
  Sort a table by IP address or subnet column
  Handles mixed IPv4 addresses and subnets
  parameters:
    table       - table to sort
    columnName  - (optional) column name or index to sort by, default is first column
    descending  - (optional) sort descending, default is false (ascending)
    prefixOrder - (optional) "none", "shortest", or "longest", default is "shortest"
                    none:    sort by IP address only, ignore prefix length
                    shortest: sort by IP address, then shortest prefix first (largest subnets)
                    longest:  sort by IP address, then longest prefix first (smallest subnets)
  example:
    IpSortTable(table, "IP", false, "shortest")
    IpSortTable(table, 0, true)  // sort first column descending, default prefixOrder
    IpSortTable(table)           // sort by first column ascending, with shortest prefix
 */
IpSortTable = (
    table as table,
    optional columnName as any,
    optional descending as logical,
    optional prefixOrder as text
) as table =>
    let
        // Determine column to sort by
        colName = if columnName is null then
                    Table.ColumnNames(table){0}
                  else if Value.Is(columnName, type number) then
                    Table.ColumnNames(table){columnName}
                  else
                    columnName,

        // Set defaults
        desc = if descending is null then false else descending,
        prefixOrd = if prefixOrder is null or prefixOrder = "" then "shortest" else prefixOrder,
        validPrefixOrders = {"none", "shortest", "longest"},
        prefixOrdValid = if List.Contains(validPrefixOrders, prefixOrd) then prefixOrd else "shortest",

        // Parse IP/subnet and add temporary columns
        withParsed = Table.AddColumn(
            table,
            "_Parsed",
            each try _ParseIpNet(Record.Field(_, colName)) otherwise null
        ),

        withIpKey = Table.AddColumn(
            withParsed,
            "_IpKey",
            each if [_Parsed] = null then null else _IpToNbr([_Parsed][Ip])
        ),

        withPrefixKey = Table.AddColumn(
            withIpKey,
            "_PrefixKey",
            each if [_Parsed] = null then null else [_Parsed][Len]
        ),

        // Add column to handle nulls (empty cells and invalid entries go to end)
        withNullFlag = Table.AddColumn(
            withPrefixKey,
            "_IsNull",
            each
                let
                    val = Record.Field(_, colName)
                in
                    if val = null or val = "" or [_Parsed] = null then 1 else 0
        ),

        // Sort based on prefix order setting
        sorted = if prefixOrdValid = "none" then
            Table.Sort(
                withNullFlag,
                {
                    {"_IsNull", Order.Ascending},
                    {"_IpKey", if desc then Order.Descending else Order.Ascending}
                }
            )
        else
            Table.Sort(
                withNullFlag,
                {
                    {"_IsNull", Order.Ascending},
                    {"_IpKey", if desc then Order.Descending else Order.Ascending},
                    {"_PrefixKey", if prefixOrdValid = "longest" then Order.Descending else Order.Ascending}
                }
            ),

        // Remove temporary columns
        result = Table.RemoveColumns(sorted, {"_Parsed", "_IpKey", "_PrefixKey", "_IsNull"})
    in
        result,

in
[
    IpIsValid = IpIsValid,
    IpStrToNbr = IpStrToNbr,
    IpBinToStr = IpBinToStr,
    IpAnd = IpAnd,
    IpOr = IpOr,
    IpXor = IpXor,
    IpWithoutMask = IpWithoutMask,
    IpSubnetLen = IpSubnetLen,
    IpMask = IpMask,
    IpWildMask = IpWildMask,
    IpInvertMask = IpInvertMask,
    IpMaskLen = IpMaskLen,
    IpAdd = IpAdd,
    IpDiff = IpDiff,
    IpGetByte = IpGetByte,
    IpSetByte = IpSetByte,
    IpSubnetSize = IpSubnetSize,
    IpClearHostBits = IpClearHostBits,
    IpIsInSubnet = IpIsInSubnet,
    IpSubnetIsInSubnet = IpSubnetIsInSubnet,
    IpSubnetMatch = IpSubnetMatch,
    IpSubnetVLookup = IpSubnetVLookup,
    IpIsPrivate = IpIsPrivate,
    IpDivideSubnet = IpDivideSubnet,
    IpRangeToCIDR = IpRangeToCIDR,
    IpSortArray = IpSortArray,
    IpSubnetSortArray = IpSubnetSortArray,
    IpFindOverlappingSubnets = IpFindOverlappingSubnets,
    IpSubnetAggregateArray = IpSubnetAggregateArray,
    IpSubtractSubnets = IpSubtractSubnets,
    IpCommonSubnets = IpCommonSubnets,
    IpMergeOnSubnet = IpMergeOnSubnet,
    IpSortTable = IpSortTable,
    Ipv6MaskLen = Ipv6MaskLen,
    Ipv6WithoutMask = Ipv6WithoutMask,
    Ipv6IsValid = Ipv6IsValid,
    Ipv6Expand = Ipv6Expand,
    Ipv6Compress = Ipv6Compress,
    Ipv6ToBin = Ipv6ToBin,
    Ipv6FromBin = Ipv6FromBin,
    Ipv6AddInt = Ipv6AddInt,
    Ipv6Add = Ipv6Add,
    Ipv6GetBlock = Ipv6GetBlock,
    Ipv6GetBlockInt = Ipv6GetBlockInt,
    Ipv6SetBlock = Ipv6SetBlock,
    Ipv6SetBlockInt = Ipv6SetBlockInt,
    Ipv6SetBits = Ipv6SetBits,
    Ipv6GetIpv4 = Ipv6GetIpv4,
    Ipv6SetIpv4 = Ipv6SetIpv4,
    Ipv6IsInSubnet = Ipv6IsInSubnet,
    Ipv6SubnetFirstAddress = Ipv6SubnetFirstAddress,
    Ipv6SubnetLastAddress = Ipv6SubnetLastAddress,
    Ipv6SortArray = Ipv6SortArray
]