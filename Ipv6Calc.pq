let
//==============================================
//   IPv4 Helper Functions (for IPv4-in-IPv6)
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

in
[
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
