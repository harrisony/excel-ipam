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
    _IpToNbr = (bytes as list) as number => List.Accumulate(bytes, 0, (state, current) => state * 256 + current),
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
                    result = {Number.Mod([value], 256)} & [result],
                    value = Number.IntegerDivide([value], 256)
                ]
            )
        in
            Generate{List.Count(Generate) - 1}[result],
/*--------------------------------------------
   _IpAnd
----------------------------------------------
 Bitwise AND on two IP byte lists
*/
    _IpAnd = (ip1 as list, ip2 as list) as list => List.Transform({0..3}, each Number.BitwiseAnd(ip1{_}, ip2{_})),
/*--------------------------------------------
   _IpOr
----------------------------------------------
 Bitwise OR on two IP byte lists
*/
    _IpOr = (ip1 as list, ip2 as list) as list => List.Transform({0..3}, each Number.BitwiseOr(ip1{_}, ip2{_})),
/*--------------------------------------------
   _IpXor
----------------------------------------------
 Bitwise XOR on two IP byte lists
*/
    _IpXor = (ip1 as list, ip2 as list) as list => List.Transform({0..3}, each Number.BitwiseXor(ip1{_}, ip2{_})),
/*--------------------------------------------
   _IpLenToMask
----------------------------------------------
 Converts a prefix length to a mask byte list
 example:
   _IpLenToMask(24) returns {255, 255, 255, 0}
*/
    _IpLenToMask = (len as number) as list =>
        let
            maskNbr = Number.BitwiseShiftLeft(Number.Power(2, len) - 1, 32 - len)
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
            parsed =
                if slashPos >= 0 then
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
                        maskBytes = List.Transform(Text.Split(maskPart, "."), Number.FromText),
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
            ipBytes = List.Transform(Text.Split(parsed[IpText], "."), Number.FromText),
            // Get mask
            mask = if Record.HasFields(parsed, "MaskBytes") then parsed[MaskBytes] else _IpLenToMask(parsed[Len]),
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
            ipText = Text.Combine(List.Transform(net[Ip], Number.ToText), "."),
            maskPart =
                if net[Cidr] then
                    "/" & Number.ToText(net[Len])
                else
                    " " & Text.Combine(List.Transform(net[Mask], Number.ToText), ".")
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
            if diff <> 0 then
                diff
            else
                a[Len] - b[Len],
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
            otherwise false,
    IpIsValidType = type function (
        ip as (
            type text meta [
                Documentation.FieldCaption = "IP Address",
                Documentation.FieldDescription = "IPv4 address to validate in dotted notation (e.g., 192.168.1.1)",
                Documentation.SampleValues = {"192.168.1.1", "10.0.0.1", "172.16.0.1"}
            ]
        )
    ) as logical meta [
        Documentation.Name = "IpIsValid",
        Documentation.LongDescription = "Returns true if an IP address is formatted correctly: no spaces, no extra zeros, no incorrect values. Validates that the address contains exactly 4 octets and each octet is properly formatted.",
        Documentation.Examples = {
            [
                Description = "Valid IP address",
                Code = "IpCalc.IpIsValid(""192.168.1.1"")",
                Result = "true"
            ],
            [
                Description = "Invalid - leading zero",
                Code = "IpCalc.IpIsValid(""192.168.001.1"")",
                Result = "false"
            ],
            [
                Description = "Invalid - missing octet",
                Code = "IpCalc.IpIsValid(""192.168.1"")",
                Result = "false"
            ]
        }
    ],
    IpIsValidDoc = Value.ReplaceType(IpIsValid, IpIsValidType),
/*--------------------------------------------
   IpStrToNbr
----------------------------------------------
 Converts a text IP address to number
 example:
   IpStrToNbr("1.2.3.4") returns 16909060
*/
    IpStrToNbr = (ip as text) as number =>
        let
            bytes = List.Transform(Text.Split(ip, "."), Number.FromText)
        in
            _IpToNbr(bytes),
    IpStrToNbrType = type function (
        ip as (
            type text meta [
                Documentation.FieldCaption = "IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation to convert to binary number",
                Documentation.SampleValues = {"1.2.3.4", "192.168.1.1", "10.0.0.1"}
            ]
        )
    ) as number meta [
        Documentation.Name = "IpStrToNbr",
        Documentation.LongDescription = "Converts a text IP address to its binary number representation. The IP address is treated as a 32-bit unsigned integer.",
        Documentation.Examples = {
            [
                Description = "Convert IP to number",
                Code = "IpCalc.IpStrToNbr(""1.2.3.4"")",
                Result = "16909060"
            ],
            [
                Description = "Convert another IP",
                Code = "IpCalc.IpStrToNbr(""192.168.1.1"")",
                Result = "3232235777"
            ]
        }
    ],
    IpStrToNbrDoc = Value.ReplaceType(IpStrToNbr, IpStrToNbrType),
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
            Text.Combine(List.Transform(bytes, Number.ToText), "."),
    IpBinToStrType = type function (
        nbr as (
            type number meta [
                Documentation.FieldCaption = "Binary Number",
                Documentation.FieldDescription = "32-bit number representing an IPv4 address",
                Documentation.SampleValues = {16909060, 3232235777, 167772161}
            ]
        )
    ) as text meta [
        Documentation.Name = "IpBinToStr",
        Documentation.LongDescription = "Converts a binary number representation to a text IP address in dotted notation.",
        Documentation.Examples = {
            [
                Description = "Convert number to IP",
                Code = "IpCalc.IpBinToStr(16909060)",
                Result = """1.2.3.4"""
            ],
            [
                Description = "Convert another number",
                Code = "IpCalc.IpBinToStr(3232235777)",
                Result = """192.168.1.1"""
            ]
        }
    ],
    IpBinToStrDoc = Value.ReplaceType(IpBinToStr, IpBinToStrType),
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
    IpAndType = type function (
        ip1 as (
            type text meta [
                Documentation.FieldCaption = "First IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation",
                Documentation.SampleValues = {"192.168.1.1", "255.255.255.0", "10.0.0.255"}
            ]
        ),
        ip2 as (
            type text meta [
                Documentation.FieldCaption = "Second IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation",
                Documentation.SampleValues = {"255.255.255.0", "192.168.1.1", "0.0.0.255"}
            ]
        )
    ) as text meta [
        Documentation.Name = "IpAnd",
        Documentation.LongDescription = "Performs a bitwise AND operation between two IP addresses. Useful for applying subnet masks or clearing host bits.",
        Documentation.Examples = {
            [
                Description = "Apply subnet mask to IP",
                Code = "IpCalc.IpAnd(""192.168.1.1"", ""255.255.255.0"")",
                Result = """192.168.1.0"""
            ],
            [
                Description = "Bitwise AND operation",
                Code = "IpCalc.IpAnd(""10.20.30.40"", ""255.0.255.0"")",
                Result = """10.0.30.0"""
            ]
        }
    ],
    IpAndDoc = Value.ReplaceType(IpAnd, IpAndType),
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
    IpOrType = type function (
        ip1 as (
            type text meta [
                Documentation.FieldCaption = "First IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation",
                Documentation.SampleValues = {"192.168.1.1", "10.0.0.0", "172.16.0.0"}
            ]
        ),
        ip2 as (
            type text meta [
                Documentation.FieldCaption = "Second IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation",
                Documentation.SampleValues = {"0.0.0.255", "192.168.1.1", "0.0.255.255"}
            ]
        )
    ) as text meta [
        Documentation.Name = "IpOr",
        Documentation.LongDescription = "Performs a bitwise OR operation between two IP addresses. Useful for setting specific bits or creating broadcast addresses.",
        Documentation.Examples = {
            [
                Description = "Create broadcast address",
                Code = "IpCalc.IpOr(""192.168.1.0"", ""0.0.0.255"")",
                Result = """192.168.1.255"""
            ],
            [
                Description = "Bitwise OR operation",
                Code = "IpCalc.IpOr(""10.0.0.0"", ""0.255.255.255"")",
                Result = """10.255.255.255"""
            ]
        }
    ],
    IpOrDoc = Value.ReplaceType(IpOr, IpOrType),
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
    IpXorType = type function (
        ip1 as (
            type text meta [
                Documentation.FieldCaption = "First IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation",
                Documentation.SampleValues = {"192.168.1.1", "10.0.0.0", "255.255.255.255"}
            ]
        ),
        ip2 as (
            type text meta [
                Documentation.FieldCaption = "Second IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation",
                Documentation.SampleValues = {"0.0.0.255", "192.168.1.1", "255.255.255.0"}
            ]
        )
    ) as text meta [
        Documentation.Name = "IpXor",
        Documentation.LongDescription = "Performs a bitwise XOR (exclusive OR) operation between two IP addresses. Each bit in the result is set if the corresponding bits in the two inputs are different.",
        Documentation.Examples = {
            [
                Description = "XOR two addresses",
                Code = "IpCalc.IpXor(""192.168.1.1"", ""0.0.0.255"")",
                Result = """192.168.1.254"""
            ],
            [
                Description = "Invert specific bits",
                Code = "IpCalc.IpXor(""255.255.255.255"", ""0.0.0.255"")",
                Result = """255.255.255.0"""
            ]
        }
    ],
    IpXorDoc = Value.ReplaceType(IpXor, IpXorType),
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
            result = if slashPos >= 0 then Text.Start(s, slashPos) else if spacePos >= 0 then
                Text.Start(s, spacePos)
            else
                s
        in
            result,
    IpWithoutMaskType = type function (
        subnet as (
            type text meta [
                Documentation.FieldCaption = "IP or Subnet",
                Documentation.FieldDescription = "IP address or subnet in CIDR (e.g., 192.168.1.0/24) or dotted mask notation (e.g., 192.168.1.0 255.255.255.0)",
                Documentation.SampleValues = {"192.168.1.0/24", "192.168.1.0 255.255.255.0", "10.0.0.1"}
            ]
        )
    ) as text meta [
        Documentation.Name = "IpWithoutMask",
        Documentation.LongDescription = "Removes the netmask notation from an IP address or subnet, returning only the IP address portion.",
        Documentation.Examples = {
            [
                Description = "Remove CIDR notation",
                Code = "IpCalc.IpWithoutMask(""192.168.1.0/24"")",
                Result = """192.168.1.0"""
            ],
            [
                Description = "Remove dotted mask notation",
                Code = "IpCalc.IpWithoutMask(""192.168.1.0 255.255.255.0"")",
                Result = """192.168.1.0"""
            ],
            [
                Description = "No mask present",
                Code = "IpCalc.IpWithoutMask(""192.168.1.1"")",
                Result = """192.168.1.1"""
            ]
        }
    ],
    IpWithoutMaskDoc = Value.ReplaceType(IpWithoutMask, IpWithoutMaskType),
/*--------------------------------------------
   IpSubnetLen
----------------------------------------------
 Get the mask length from a subnet
 example:
   IpSubnetLen("192.168.1.1/24") returns 24
   IpSubnetLen("192.168.1.1 255.255.255.0") returns 24
   IpSubnetLen("192.168.1.1") returns 32
*/
    IpSubnetLen = (s as text) as number => let net = _ParseIpNet(s) in net[Len],
    IpSubnetLenType = type function (
        subnet as (
            type text meta [
                Documentation.FieldCaption = "Subnet",
                Documentation.FieldDescription = "Subnet in CIDR or dotted mask notation. If no mask is specified, returns 32 (host address).",
                Documentation.SampleValues = {"192.168.1.0/24", "192.168.1.0 255.255.255.0", "10.0.0.0/8"}
            ]
        )
    ) as number meta [
        Documentation.Name = "IpSubnetLen",
        Documentation.LongDescription = "Returns the prefix length (subnet mask length) from a subnet. Supports both CIDR and dotted mask notation. Returns 32 for plain IP addresses without a mask.",
        Documentation.Examples = {
            [
                Description = "Get prefix from CIDR",
                Code = "IpCalc.IpSubnetLen(""192.168.1.0/24"")",
                Result = "24"
            ],
            [
                Description = "Get prefix from dotted mask",
                Code = "IpCalc.IpSubnetLen(""192.168.1.0 255.255.255.0"")",
                Result = "24"
            ],
            [
                Description = "Plain IP address",
                Code = "IpCalc.IpSubnetLen(""192.168.1.1"")",
                Result = "32"
            ]
        }
    ],
    IpSubnetLenDoc = Value.ReplaceType(IpSubnetLen, IpSubnetLenType),
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
            net = _ParseIpNet(s), maskText = Text.Combine(List.Transform(net[Mask], Number.ToText), ".")
        in
            maskText,
    IpMaskType = type function (
        subnet as (
            type text meta [
                Documentation.FieldCaption = "Subnet",
                Documentation.FieldDescription = "Subnet in CIDR or dotted mask notation",
                Documentation.SampleValues = {"192.168.1.0/24", "10.0.0.0/8", "172.16.0.0 255.255.0.0"}
            ]
        )
    ) as text meta [
        Documentation.Name = "IpMask",
        Documentation.LongDescription = "Returns the subnet mask in dotted notation from a subnet. Accepts both CIDR and dotted mask notation as input.",
        Documentation.Examples = {
            [
                Description = "Get mask from CIDR",
                Code = "IpCalc.IpMask(""192.168.1.0/24"")",
                Result = """255.255.255.0"""
            ],
            [
                Description = "Get mask from /8 network",
                Code = "IpCalc.IpMask(""10.0.0.0/8"")",
                Result = """255.0.0.0"""
            ],
            [
                Description = "Get mask from dotted notation",
                Code = "IpCalc.IpMask(""172.16.0.0 255.255.0.0"")",
                Result = """255.255.0.0"""
            ]
        }
    ],
    IpMaskDoc = Value.ReplaceType(IpMask, IpMaskType),
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
            net = _ParseIpNet(s), wildNbr = Number.Power(2, 32 - net[Len]) - 1, wildBytes = _IpFromNbr(wildNbr)
        in
            Text.Combine(List.Transform(wildBytes, Number.ToText), "."),
    IpWildMaskType = type function (
        subnet as (
            type text meta [
                Documentation.FieldCaption = "Subnet",
                Documentation.FieldDescription = "Subnet in CIDR or dotted mask notation",
                Documentation.SampleValues = {"192.168.1.0/24", "10.0.0.0/8", "172.16.0.0 255.255.0.0"}
            ]
        )
    ) as text meta [
        Documentation.Name = "IpWildMask",
        Documentation.LongDescription = "Returns the wildcard (inverse) mask from a subnet. A wildcard mask is the bitwise inverse of the subnet mask, commonly used in ACLs and routing configurations.",
        Documentation.Examples = {
            [
                Description = "Get wildcard mask from /24",
                Code = "IpCalc.IpWildMask(""192.168.1.0/24"")",
                Result = """0.0.0.255"""
            ],
            [
                Description = "Get wildcard from /8 network",
                Code = "IpCalc.IpWildMask(""10.0.0.0/8"")",
                Result = """0.255.255.255"""
            ],
            [
                Description = "Get wildcard from dotted notation",
                Code = "IpCalc.IpWildMask(""172.16.0.0 255.255.0.0"")",
                Result = """0.0.255.255"""
            ]
        }
    ],
    IpWildMaskDoc = Value.ReplaceType(IpWildMask, IpWildMaskType),
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
            maskNbr = IpStrToNbr(mask), invertedNbr = Number.Power(2, 32) - 1 - maskNbr
        in
            IpBinToStr(invertedNbr),
    IpInvertMaskType = type function (
        mask as (
            type text meta [
                Documentation.FieldCaption = "Mask",
                Documentation.FieldDescription = "Subnet mask or wildcard mask in dotted notation",
                Documentation.SampleValues = {"255.255.255.0", "0.0.0.255", "255.255.0.0"}
            ]
        )
    ) as text meta [
        Documentation.Name = "IpInvertMask",
        Documentation.LongDescription = "Inverts a subnet mask to a wildcard mask, or a wildcard mask to a subnet mask. Performs a bitwise NOT operation on the mask.",
        Documentation.Examples = {
            [
                Description = "Convert subnet mask to wildcard",
                Code = "IpCalc.IpInvertMask(""255.255.255.0"")",
                Result = """0.0.0.255"""
            ],
            [
                Description = "Convert wildcard to subnet mask",
                Code = "IpCalc.IpInvertMask(""0.0.0.255"")",
                Result = """255.255.255.0"""
            ],
            [
                Description = "Invert /16 mask",
                Code = "IpCalc.IpInvertMask(""255.255.0.0"")",
                Result = """0.0.255.255"""
            ]
        }
    ],
    IpInvertMaskDoc = Value.ReplaceType(IpInvertMask, IpInvertMaskType),
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
                [
                    count = 0,
                    remaining = notMask
                ],
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
    IpMaskLenType = type function (
        mask as (
            type text meta [
                Documentation.FieldCaption = "Subnet Mask",
                Documentation.FieldDescription = "Subnet mask in dotted notation (e.g., 255.255.255.0)",
                Documentation.SampleValues = {"255.255.255.0", "255.0.0.0", "255.255.0.0"}
            ]
        )
    ) as number meta [
        Documentation.Name = "IpMaskLen",
        Documentation.LongDescription = "Returns the prefix length from a subnet mask given in dotted notation. Converts a mask like 255.255.255.0 to prefix length 24.",
        Documentation.Examples = {
            [
                Description = "Get prefix from /24 mask",
                Code = "IpCalc.IpMaskLen(""255.255.255.0"")",
                Result = "24"
            ],
            [
                Description = "Get prefix from /8 mask",
                Code = "IpCalc.IpMaskLen(""255.0.0.0"")",
                Result = "8"
            ],
            [
                Description = "Get prefix from /16 mask",
                Code = "IpCalc.IpMaskLen(""255.255.0.0"")",
                Result = "16"
            ]
        }
    ],
    IpMaskLenDoc = Value.ReplaceType(IpMaskLen, IpMaskLenType),
/*--------------------------------------------
   IpAdd
----------------------------------------------
 Add offset to IP address
 example:
   IpAdd("192.168.1.1", 4) returns "192.168.1.5"
   IpAdd("192.168.1.1", 256) returns "192.168.2.1"
*/
    IpAdd = (ip as text, offset as number) as text => IpBinToStr(IpStrToNbr(ip) + offset),
    IpAddType = type function (
        ip as (
            type text meta [
                Documentation.FieldCaption = "IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation",
                Documentation.SampleValues = {"192.168.1.1", "10.0.0.1", "172.16.0.0"}
            ]
        ),
        offset as (
            type number meta [
                Documentation.FieldCaption = "Offset",
                Documentation.FieldDescription = "Number to add to the IP address (can be positive or negative). Handles overflow across octets.",
                Documentation.SampleValues = {1, 4, 256, -10}
            ]
        )
    ) as text meta [
        Documentation.Name = "IpAdd",
        Documentation.LongDescription = "Adds a numeric offset to an IP address. The offset is applied to the 32-bit representation of the IP address, properly handling overflow across octets.",
        Documentation.Examples = {
            [
                Description = "Add within same octet",
                Code = "IpCalc.IpAdd(""192.168.1.1"", 4)",
                Result = """192.168.1.5"""
            ],
            [
                Description = "Add across octets",
                Code = "IpCalc.IpAdd(""192.168.1.1"", 256)",
                Result = """192.168.2.1"""
            ],
            [
                Description = "Subtract using negative offset",
                Code = "IpCalc.IpAdd(""192.168.1.10"", -5)",
                Result = """192.168.1.5"""
            ]
        }
    ],
    IpAddDoc = Value.ReplaceType(IpAdd, IpAddType),
/*--------------------------------------------
   IpDiff
----------------------------------------------
 Calculate difference between two IP addresses
 example:
   IpDiff("192.168.1.7", "192.168.1.1") returns 6
*/
    IpDiff = (ip1 as text, ip2 as text) as number => IpStrToNbr(ip1) - IpStrToNbr(ip2),
    IpDiffType = type function (
        ip1 as (
            type text meta [
                Documentation.FieldCaption = "First IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation",
                Documentation.SampleValues = {"192.168.1.7", "10.0.0.100", "172.16.1.0"}
            ]
        ),
        ip2 as (
            type text meta [
                Documentation.FieldCaption = "Second IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation to subtract from the first",
                Documentation.SampleValues = {"192.168.1.1", "10.0.0.1", "172.16.0.0"}
            ]
        )
    ) as number meta [
        Documentation.Name = "IpDiff",
        Documentation.LongDescription = "Calculates the numeric difference between two IP addresses. Returns the number of addresses between them.",
        Documentation.Examples = {
            [
                Description = "Calculate difference",
                Code = "IpCalc.IpDiff(""192.168.1.7"", ""192.168.1.1"")",
                Result = "6"
            ],
            [
                Description = "Difference across octets",
                Code = "IpCalc.IpDiff(""192.168.2.1"", ""192.168.1.1"")",
                Result = "256"
            ],
            [
                Description = "Same addresses",
                Code = "IpCalc.IpDiff(""10.0.0.1"", ""10.0.0.1"")",
                Result = "0"
            ]
        }
    ],
    IpDiffDoc = Value.ReplaceType(IpDiff, IpDiffType),
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
    IpGetByteType = type function (
        ip as (
            type text meta [
                Documentation.FieldCaption = "IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation",
                Documentation.SampleValues = {"192.168.1.1", "10.20.30.40", "172.16.0.1"}
            ]
        ),
        pos as (
            type number meta [
                Documentation.FieldCaption = "Position",
                Documentation.FieldDescription = "Octet position (1-4): 1=first octet, 4=last octet",
                Documentation.SampleValues = {1, 2, 3, 4}
            ]
        )
    ) as number meta [
        Documentation.Name = "IpGetByte",
        Documentation.LongDescription = "Returns one octet (byte) from an IP address at the specified position. Position 1 is the first octet, position 4 is the last.",
        Documentation.Examples = {
            [
                Description = "Get first octet",
                Code = "IpCalc.IpGetByte(""192.168.1.1"", 1)",
                Result = "192"
            ],
            [
                Description = "Get last octet",
                Code = "IpCalc.IpGetByte(""192.168.1.1"", 4)",
                Result = "1"
            ],
            [
                Description = "Get third octet",
                Code = "IpCalc.IpGetByte(""10.20.30.40"", 3)",
                Result = "30"
            ]
        }
    ],
    IpGetByteDoc = Value.ReplaceType(IpGetByte, IpGetByteType),
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
    IpSetByteType = type function (
        ip as (
            type text meta [
                Documentation.FieldCaption = "IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation",
                Documentation.SampleValues = {"192.168.1.1", "10.20.30.40", "172.16.0.1"}
            ]
        ),
        pos as (
            type number meta [
                Documentation.FieldCaption = "Position",
                Documentation.FieldDescription = "Octet position to modify (1-4): 1=first octet, 4=last octet",
                Documentation.SampleValues = {1, 2, 3, 4}
            ]
        ),
        newvalue as (
            type number meta [
                Documentation.FieldCaption = "New Value",
                Documentation.FieldDescription = "New value for the octet (0-255)",
                Documentation.SampleValues = {0, 1, 10, 20, 255}
            ]
        )
    ) as text meta [
        Documentation.Name = "IpSetByte",
        Documentation.LongDescription = "Sets one octet (byte) in an IP address to a new value. Position 1 is the first octet, position 4 is the last.",
        Documentation.Examples = {
            [
                Description = "Set last octet",
                Code = "IpCalc.IpSetByte(""192.168.1.1"", 4, 20)",
                Result = """192.168.1.20"""
            ],
            [
                Description = "Set first octet",
                Code = "IpCalc.IpSetByte(""192.168.1.1"", 1, 10)",
                Result = """10.168.1.1"""
            ],
            [
                Description = "Set third octet",
                Code = "IpCalc.IpSetByte(""10.20.30.40"", 3, 100)",
                Result = """10.20.100.40"""
            ]
        }
    ],
    IpSetByteDoc = Value.ReplaceType(IpSetByte, IpSetByteType),
/*--------------------------------------------
   IpSubnetSize
----------------------------------------------
 Returns the number of addresses in a subnet
 example:
   IpSubnetSize("192.168.1.32/29") returns 8
   IpSubnetSize("192.168.1.0 255.255.255.0") returns 256
*/
    IpSubnetSize = (subnet as text) as number => let net = _ParseIpNet(subnet) in Number.Power(2, 32 - net[Len]),
    IpSubnetSizeType = type function (
        subnet as (
            type text meta [
                Documentation.FieldCaption = "Subnet",
                Documentation.FieldDescription = "Subnet in CIDR or dotted mask notation",
                Documentation.SampleValues = {"192.168.1.0/24", "10.0.0.0/8", "192.168.1.32/29"}
            ]
        )
    ) as number meta [
        Documentation.Name = "IpSubnetSize",
        Documentation.LongDescription = "Returns the total number of IP addresses in a subnet. Supports both CIDR and dotted mask notation.",
        Documentation.Examples = {
            [
                Description = "Size of /24 network",
                Code = "IpCalc.IpSubnetSize(""192.168.1.0/24"")",
                Result = "256"
            ],
            [
                Description = "Size of /29 network",
                Code = "IpCalc.IpSubnetSize(""192.168.1.32/29"")",
                Result = "8"
            ],
            [
                Description = "Size with dotted mask",
                Code = "IpCalc.IpSubnetSize(""10.0.0.0 255.0.0.0"")",
                Result = "16777216"
            ]
        }
    ],
    IpSubnetSizeDoc = Value.ReplaceType(IpSubnetSize, IpSubnetSizeType),
/*--------------------------------------------
   IpClearHostBits
----------------------------------------------
 Set to zero the bits in the host part of an address
 example:
   IpClearHostBits("192.168.1.1/24") returns "192.168.1.0/24"
   IpClearHostBits("192.168.1.193 255.255.255.128") returns "192.168.1.128 255.255.255.128"
*/
    IpClearHostBits = (net as text) as text => _IpNetToString(_ParseIpNet(net)),
    IpClearHostBitsType = type function (
        subnet as (
            type text meta [
                Documentation.FieldCaption = "Subnet",
                Documentation.FieldDescription = "Subnet in CIDR or dotted mask notation",
                Documentation.SampleValues = {"192.168.1.1/24", "10.0.0.5 255.0.0.0", "172.16.1.193 255.255.255.128"}
            ]
        )
    ) as text meta [
        Documentation.Name = "IpClearHostBits",
        Documentation.LongDescription = "Sets the host bits to zero in a subnet, returning the network address. Preserves the original notation format (CIDR or dotted mask).",
        Documentation.Examples = {
            [
                Description = "Clear host bits in CIDR notation",
                Code = "IpCalc.IpClearHostBits(""192.168.1.1/24"")",
                Result = """192.168.1.0/24"""
            ],
            [
                Description = "Clear host bits in dotted mask",
                Code = "IpCalc.IpClearHostBits(""192.168.1.193 255.255.255.128"")",
                Result = """192.168.1.128 255.255.255.128"""
            ]
        }
    ],
    IpClearHostBitsDoc = Value.ReplaceType(IpClearHostBits, IpClearHostBitsType),
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
    IpIsInSubnetType = type function (
        ip as (
            type text meta [
                Documentation.FieldCaption = "IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation",
                Documentation.SampleValues = {"192.168.1.35", "10.0.0.1", "172.16.5.100"}
            ]
        ),
        subnet as (
            type text meta [
                Documentation.FieldCaption = "Subnet",
                Documentation.FieldDescription = "Subnet in CIDR or dotted mask notation to check containment",
                Documentation.SampleValues = {"192.168.1.0/24", "192.168.1.32/29", "10.0.0.0/8"}
            ]
        )
    ) as logical meta [
        Documentation.Name = "IpIsInSubnet",
        Documentation.LongDescription = "Returns true if the IP address falls within the specified subnet range. Supports both CIDR and dotted mask notation.",
        Documentation.Examples = {
            [
                Description = "IP is in subnet",
                Code = "IpCalc.IpIsInSubnet(""192.168.1.35"", ""192.168.1.32/29"")",
                Result = "true"
            ],
            [
                Description = "IP is not in subnet",
                Code = "IpCalc.IpIsInSubnet(""192.168.1.41"", ""192.168.1.32/29"")",
                Result = "false"
            ]
        }
    ],
    IpIsInSubnetDoc = Value.ReplaceType(IpIsInSubnet, IpIsInSubnetType),
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
            net1 = _ParseIpNet(subnet1), net2 = _ParseIpNet(subnet2)
        in
            if net1[Len] < net2[Len] then
                false
            else
                net1[Ip] = _IpAnd(net1[Ip], net2[Mask]),
    IpSubnetIsInSubnetType = type function (
        subnet1 as (
            type text meta [
                Documentation.FieldCaption = "Subnet to Test",
                Documentation.FieldDescription = "Subnet to check if it's contained in another subnet",
                Documentation.SampleValues = {"192.168.1.32/29", "10.1.0.0/16", "172.16.5.0/24"}
            ]
        ),
        subnet2 as (
            type text meta [
                Documentation.FieldCaption = "Containing Subnet",
                Documentation.FieldDescription = "Subnet to check if it contains the first subnet",
                Documentation.SampleValues = {"192.168.1.0/24", "10.0.0.0/8", "172.16.0.0/16"}
            ]
        )
    ) as logical meta [
        Documentation.Name = "IpSubnetIsInSubnet",
        Documentation.LongDescription = "Returns true if subnet1 is completely contained within subnet2. Both subnets must be in CIDR or dotted mask notation.",
        Documentation.Examples = {
            [
                Description = "Subnet is contained",
                Code = "IpCalc.IpSubnetIsInSubnet(""192.168.1.32/29"", ""192.168.1.0/24"")",
                Result = "true"
            ],
            [
                Description = "Subnet is not contained",
                Code = "IpCalc.IpSubnetIsInSubnet(""192.168.1.32/28"", ""192.168.1.32/29"")",
                Result = "false"
            ]
        }
    ],
    IpSubnetIsInSubnetDoc = Value.ReplaceType(IpSubnetIsInSubnet, IpSubnetIsInSubnetType),
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
                [
                    bestRow = 0,
                    bestLen = -1
                ],
                (state, idx) =>
                    let
                        subnetText = if List.Count(table{idx}) > 0 then table{idx}{0} else table{idx},
                        net = _ParseIpNet(subnetText)
                    in
                        if
                            net[Len] > state[bestLen]
                            and searchNet[Len] >= net[Len]
                            and _IpAnd(searchNet[Ip], net[Mask]) = net[Ip]
                        then
                            [bestRow = idx + 1, bestLen = net[Len]]
                        else
                            state
            )
        in
            FindBest[bestRow],
    IpSubnetMatchType = type function (
        ip as (
            type text meta [
                Documentation.FieldCaption = "IP or Subnet",
                Documentation.FieldDescription = "IP address or subnet to match against the list",
                Documentation.SampleValues = {"192.168.1.35", "10.1.5.100"}
            ]
        ),
        table as (
            type list meta [
                Documentation.FieldCaption = "Subnet List",
                Documentation.FieldDescription = "List of subnets to search for best match",
                Documentation.SampleValues = {{{"192.168.1.0/24"}, {"192.168.1.32/29"}}}
            ]
        )
    ) as number meta [
        Documentation.Name = "IpSubnetMatch",
        Documentation.LongDescription = "Searches for the best matching subnet (most specific/longest prefix) that contains the IP address. Returns the 1-based row number of the match, or 0 if no match found.",
        Documentation.Examples = {
            [
                Description = "Find best match (most specific)",
                Code = "IpCalc.IpSubnetMatch(""192.168.1.35"", {{""192.168.1.0/24""}, {""192.168.1.32/29""}})",
                Result = "2"
            ]
        }
    ],
    IpSubnetMatchDoc = Value.ReplaceType(IpSubnetMatch, IpSubnetMatchType),
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
    IpSubnetVLookupType = type function (
        ip as (
            type text meta [
                Documentation.FieldCaption = "IP or Subnet",
                Documentation.FieldDescription = "IP address or subnet to match against the list",
                Documentation.SampleValues = {"192.168.1.35", "10.1.5.100"}
            ]
        ),
        table as (
            type list meta [
                Documentation.FieldCaption = "Lookup Table",
                Documentation.FieldDescription = "List of rows where first column contains subnets and other columns contain associated data",
                Documentation.SampleValues = {{{"192.168.1.0/24", "LAN"}, {"10.0.0.0/8", "Private"}}}
            ]
        ),
        index as (
            type number meta [
                Documentation.FieldCaption = "Column Index",
                Documentation.FieldDescription = "1-based column index to return from matched row",
                Documentation.SampleValues = {1, 2, 3}
            ]
        )
    ) as any meta [
        Documentation.Name = "IpSubnetVLookup",
        Documentation.LongDescription = "Finds the best matching subnet (most specific) for an IP address and returns a value from the specified column. Similar to Excel's VLOOKUP but with subnet matching. Returns 'Not Found' if no match.",
        Documentation.Examples = {
            [
                Description = "Lookup network name",
                Code = "IpCalc.IpSubnetVLookup(""192.168.1.35"", {{""192.168.1.0/24"", ""LAN""}, {""10.0.0.0/8"", ""Private""}}, 2)",
                Result = """LAN"""
            ]
        }
    ],
    IpSubnetVLookupDoc = Value.ReplaceType(IpSubnetVLookup, IpSubnetVLookupType),
/*--------------------------------------------
   IpIsPrivate
----------------------------------------------
 Returns TRUE if "ip" is in one of the private IP address ranges
 example:
   IpIsPrivate("192.168.1.35") returns true
   IpIsPrivate("209.85.148.104") returns false
*/
    IpIsPrivate = (ip as text) as logical =>
        IpIsInSubnet(ip, "10.0.0.0/8") or IpIsInSubnet(ip, "172.16.0.0/12") or IpIsInSubnet(ip, "192.168.0.0/16"),
    IpIsPrivateType = type function (
        ip as (
            type text meta [
                Documentation.FieldCaption = "IP Address",
                Documentation.FieldDescription = "IPv4 address in dotted notation",
                Documentation.SampleValues = {"192.168.1.35", "10.0.0.1", "8.8.8.8"}
            ]
        )
    ) as logical meta [
        Documentation.Name = "IpIsPrivate",
        Documentation.LongDescription = "Returns true if the IP address is in one of the private address ranges: 10.0.0.0/8, 172.16.0.0/12, or 192.168.0.0/16.",
        Documentation.Examples = {
            [
                Description = "Private IP (192.168.x.x)",
                Code = "IpCalc.IpIsPrivate(""192.168.1.35"")",
                Result = "true"
            ],
            [
                Description = "Public IP",
                Code = "IpCalc.IpIsPrivate(""8.8.8.8"")",
                Result = "false"
            ]
        }
    ],
    IpIsPrivateDoc = Value.ReplaceType(IpIsPrivate, IpIsPrivateType),
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
            result =
                if newLen > 32 then
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
                                    subnetText =
                                        if net[Cidr] then
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
    IpDivideSubnetType = type function (
        subnet as (
            type text meta [
                Documentation.FieldCaption = "Subnet",
                Documentation.FieldDescription = "Subnet to divide in CIDR or dotted mask notation",
                Documentation.SampleValues = {"192.168.1.0/24", "10.0.0.0/8"}
            ]
        ),
        n as (
            type number meta [
                Documentation.FieldCaption = "Bits to Add",
                Documentation.FieldDescription = "Number of bits to add to the subnet prefix (divides into 2^n smaller subnets)",
                Documentation.SampleValues = {1, 2, 3, 4}
            ]
        )
    ) as list meta [
        Documentation.Name = "IpDivideSubnet",
        Documentation.LongDescription = "Divides a subnet into smaller subnets by adding bits to the prefix length. Adding n bits creates 2^n smaller subnets. Preserves original notation format (CIDR or dotted mask).",
        Documentation.Examples = {
            [
                Description = "Divide /24 into 4 subnets (add 2 bits)",
                Code = "IpCalc.IpDivideSubnet(""192.168.1.0/24"", 2)",
                Result = "{""192.168.1.0/26"", ""192.168.1.64/26"", ""192.168.1.128/26"", ""192.168.1.192/26""}"
            ]
        }
    ],
    IpDivideSubnetDoc = Value.ReplaceType(IpDivideSubnet, IpDivideSubnetType),
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
                        [
                            len = 32,
                            ip1 = currentFirst,
                            ip2 = currentFirst,
                            diff = IpDiff(targetLast, currentFirst)
                        ],
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
    IpRangeToCIDRType = type function (
        firstAddr as (
            type text meta [
                Documentation.FieldCaption = "First IP Address",
                Documentation.FieldDescription = "Starting IP address of the range in dotted notation",
                Documentation.SampleValues = {"10.0.0.1", "192.168.1.1"}
            ]
        ),
        lastAddr as (
            type text meta [
                Documentation.FieldCaption = "Last IP Address",
                Documentation.FieldDescription = "Ending IP address of the range in dotted notation",
                Documentation.SampleValues = {"10.0.0.254", "192.168.1.255"}
            ]
        )
    ) as list meta [
        Documentation.Name = "IpRangeToCIDR",
        Documentation.LongDescription = "Converts an IP address range into a list of CIDR subnets that cover the range. Returns the minimum number of subnets needed to represent the range.",
        Documentation.Examples = {
            [
                Description = "Range that fits in one subnet",
                Code = "IpCalc.IpRangeToCIDR(""10.0.0.1"", ""10.0.0.254"")",
                Result = "{""10.0.0.0/24""}"
            ],
            [
                Description = "Range requiring multiple subnets",
                Code = "IpCalc.IpRangeToCIDR(""10.0.0.1"", ""10.0.1.63"")",
                Result = "{""10.0.0.0/24"", ""10.0.1.0/26""}"
            ]
        }
    ],
    IpRangeToCIDRDoc = Value.ReplaceType(IpRangeToCIDR, IpRangeToCIDRType),
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
                        aNbr = IpStrToNbr(a), bNbr = IpStrToNbr(b)
                    in
                        if aNbr < bNbr then
                            -1
                        else if aNbr > bNbr then
                            1
                        else
                            0
            )
        in
            sorted,
    IpSortArrayType = type function (
        arr as (
            type list meta [
                Documentation.FieldCaption = "IP Array",
                Documentation.FieldDescription = "List of IP addresses to sort",
                Documentation.SampleValues = {{"192.168.1.10", "10.0.0.1", "192.168.1.5"}}
            ]
        )
    ) as list meta [
        Documentation.Name = "IpSortArray",
        Documentation.LongDescription = "Sorts a list of IP addresses in ascending order. Empty values and null entries are filtered out.",
        Documentation.Examples = {
            [
                Description = "Sort IP addresses",
                Code = "IpCalc.IpSortArray({""192.168.1.10"", ""10.0.0.1"", ""192.168.1.5""})",
                Result = "{""10.0.0.1"", ""192.168.1.5"", ""192.168.1.10""}"
            ]
        }
    ],
    IpSortArrayDoc = Value.ReplaceType(IpSortArray, IpSortArrayType),
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
    IpSubnetSortArrayType = type function (
        arr as (
            type list meta [
                Documentation.FieldCaption = "Subnet Array",
                Documentation.FieldDescription = "List of subnets to sort (CIDR or dotted mask notation)",
                Documentation.SampleValues = {{"192.168.1.0/24", "10.0.0.0/8", "192.168.1.32/29"}}
            ]
        )
    ) as list meta [
        Documentation.Name = "IpSubnetSortArray",
        Documentation.LongDescription = "Sorts a list of subnets by network address. When network addresses are equal, shorter prefixes (larger subnets) come first. Empty values and null entries are filtered out.",
        Documentation.Examples = {
            [
                Description = "Sort subnets by network address",
                Code = "IpCalc.IpSubnetSortArray({""192.168.1.0/24"", ""10.0.0.0/8"", ""192.168.1.32/29""})",
                Result = "{""10.0.0.0/8"", ""192.168.1.0/24"", ""192.168.1.32/29""}"
            ]
        }
    ],
    IpSubnetSortArrayDoc = Value.ReplaceType(IpSubnetSortArray, IpSubnetSortArrayType),
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
                                        i <> j and netI[Len] > netJ[Len] and _IpAnd(netI[Ip], netJ[Mask]) = netJ[Ip]
                            ),
                            null
                        )
                    in
                        if containing = null then
                            ""
                        else
                            _IpNetToString(nets{containing})
            )
        in
            results,
    IpFindOverlappingSubnetsType = type function (
        arr as (
            type list meta [
                Documentation.FieldCaption = "Subnet Array",
                Documentation.FieldDescription = "List of subnets to check for overlaps",
                Documentation.SampleValues = {{"192.168.1.0/24", "192.168.1.32/29", "10.0.0.0/8"}}
            ]
        )
    ) as list meta [
        Documentation.Name = "IpFindOverlappingSubnets",
        Documentation.LongDescription = "Finds overlapping subnets in a list. For each subnet, returns the larger subnet that contains it, or empty string if no overlap exists.",
        Documentation.Examples = {
            [
                Description = "Find overlaps in subnet list",
                Code = "IpCalc.IpFindOverlappingSubnets({""192.168.1.0/24"", ""192.168.1.32/29"", ""10.0.0.0/8""})",
                Result = "{"""", ""192.168.1.0/24"", """"}"
            ]
        }
    ],
    IpFindOverlappingSubnetsDoc = Value.ReplaceType(IpFindOverlappingSubnets, IpFindOverlappingSubnetsType),
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
                        nextIncluded = next[Len] >= current[Len] and _IpAnd(next[Ip], current[Mask]) = current[Ip],
                        // Check if can join (same length and contiguous)
                        canJoin =
                            if not nextIncluded and current[Len] = next[Len] and current[Len] > 0 then
                                let
                                    // Try creating a subnet with mask 1 bit shorter
                                    bignetLen = current[Len] - 1,
                                    bignetMask = _IpLenToMask(bignetLen),
                                    bignetIp = _IpAnd(current[Ip], bignetMask)
                                in
                                    _IpAnd(next[Ip], bignetMask) = bignetIp
                            else
                                false,
                        newList =
                            if nextIncluded then
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
                                    List.ReplaceRange(List.RemoveRange(netList, index + 1, 1), index, 1, {bignet})
                            else
                                netList,
                        nextIndex = if nextIncluded or canJoin then
                        // Stay at current index to check if we can aggregate more
                        if index > 0 then index - 1 else index else index + 1
                    in
                        @Aggregate(newList, nextIndex),
            aggregated = Aggregate(sorted, 0),
            // Convert back to strings
            result = List.Transform(aggregated, _IpNetToString)
        in
            result,
    IpSubnetAggregateArrayType = type function (
        arr as (
            type list meta [
                Documentation.FieldCaption = "Subnet Array",
                Documentation.FieldDescription = "List of subnets to aggregate",
                Documentation.SampleValues = {{"192.168.1.0/25", "192.168.1.128/25", "192.168.1.32/29"}}
            ]
        )
    ) as list meta [
        Documentation.Name = "IpSubnetAggregateArray",
        Documentation.LongDescription = "Sorts and aggregates subnets by: (1) removing subnets contained in larger subnets, (2) joining contiguous subnets of the same size, (3) removing duplicates. Returns the minimal set of subnets covering the same address space.",
        Documentation.Examples = {
            [
                Description = "Aggregate contiguous /25 subnets into /24",
                Code = "IpCalc.IpSubnetAggregateArray({""192.168.1.0/25"", ""192.168.1.128/25""})",
                Result = "{""192.168.1.0/24""}"
            ]
        }
    ],
    IpSubnetAggregateArrayDoc = Value.ReplaceType(IpSubnetAggregateArray, IpSubnetAggregateArrayType),
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
                        subContainsIn = inNet[Len] >= subNet[Len] and _IpAnd(inNet[Ip], subNet[Mask]) = subNet[Ip],
                        // Check if input net contains subtract net
                        inContainsSub = subNet[Len] > inNet[Len] and _IpAnd(subNet[Ip], inNet[Mask]) = inNet[Ip],
                        newState =
                            if subContainsIn then
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
                                    inNbr = _IpToNbr(inNet[Ip]), subNbr = _IpToNbr(subNet[Ip])
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
    IpSubtractSubnetsType = type function (
        inputArray as (
            type list meta [
                Documentation.FieldCaption = "Input Subnets",
                Documentation.FieldDescription = "List of subnets representing the address space to start with",
                Documentation.SampleValues = {{"192.168.1.0/24"}}
            ]
        ),
        subtractArray as (
            type list meta [
                Documentation.FieldCaption = "Subnets to Remove",
                Documentation.FieldDescription = "List of subnets to remove from the input address space",
                Documentation.SampleValues = {{"192.168.1.0/25"}}
            ]
        )
    ) as list meta [
        Documentation.Name = "IpSubtractSubnets",
        Documentation.LongDescription = "Subtracts one set of subnets from another, returning the remaining address space. Useful for finding available IP ranges after allocations.",
        Documentation.Examples = {
            [
                Description = "Subtract half of a /24",
                Code = "IpCalc.IpSubtractSubnets({""192.168.1.0/24""}, {""192.168.1.0/25""})",
                Result = "{""192.168.1.128/25""}"
            ]
        }
    ],
    IpSubtractSubnetsDoc = Value.ReplaceType(IpSubtractSubnets, IpSubtractSubnetsType),
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
    IpCommonSubnetsType = type function (
        inputArray1 as (
            type list meta [
                Documentation.FieldCaption = "First Subnet List",
                Documentation.FieldDescription = "First list of subnets",
                Documentation.SampleValues = {{"192.168.0.0/23"}}
            ]
        ),
        inputArray2 as (
            type list meta [
                Documentation.FieldCaption = "Second Subnet List",
                Documentation.FieldDescription = "Second list of subnets",
                Documentation.SampleValues = {{"192.168.1.0/24"}}
            ]
        )
    ) as list meta [
        Documentation.Name = "IpCommonSubnets",
        Documentation.LongDescription = "Returns the IP address ranges that are common to both input subnet lists. Finds the intersection of two sets of subnets.",
        Documentation.Examples = {
            [
                Description = "Find common address space",
                Code = "IpCalc.IpCommonSubnets({""192.168.0.0/23""}, {""192.168.1.0/24""})",
                Result = "{""192.168.1.0/24""}"
            ]
        }
    ],
    IpCommonSubnetsDoc = Value.ReplaceType(IpCommonSubnets, IpCommonSubnetsType),
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
            result =
                if not isValidJoinKind then
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
                                    matchIndex =
                                        if ipValue = null or ipValue = "" then
                                            null
                                        else
                                            try
                                                let
                                                    ipBytes = List.Transform(
                                                        Text.Split(ipValue, "."), Number.FromText
                                                    ),
                                                    // Find best match (most specific subnet)
                                                    FindBest = List.Accumulate(
                                                        Table.ToRecords(rightWithIndex),
                                                        [
                                                            bestIndex = null,
                                                            bestLen = -1
                                                        ],
                                                        (state, rightRow) =>
                                                            let
                                                                subnetValue = Record.Field(
                                                                    rightRow, rightSubnetColumn
                                                                )
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
                                                                            if isMatch and net[Len] > state[bestLen]
                                                                            then
                                                                                [
                                                                                    bestIndex = Record.Field(
                                                                                        rightRow, "__RightIndex"
                                                                                    ),
                                                                                    bestLen = net[Len]
                                                                                ]
                                                                            else
                                                                                state
                                                                        otherwise state
                                                    )
                                                in
                                                    FindBest[bestIndex]
                                                otherwise null
                                in
                                    matchIndex,
                            type nullable number
                        ),
                        // Build result based on join kind
                        finalResult =
                            if actualJoinKind = "Inner" then
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
                                        joined, "__RightData", rightColumns, rightColumns
                                    ),
                                    cleaned = Table.RemoveColumns(
                                        expanded, {"__LeftIndex", "__RightIndex", "__MatchIndex"}
                                    )
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
                                        joined, "__RightData", rightColumns, rightColumns
                                    ),
                                    cleaned = Table.RemoveColumns(
                                        expanded, {"__LeftIndex", "__RightIndex", "__MatchIndex"}
                                    )
                                in
                                    cleaned
                            else if actualJoinKind = "Right" then
                                // Right join: all right rows, nulls for non-matches
                                let
                                    // Find which right rows have matches
                                    matchedRightIndexes = List.Distinct(
                                        List.Select(Table.Column(leftWithMatch, "__MatchIndex"), each _ <> null)
                                    ),
                                    rightWithMatchFlag = Table.AddColumn(
                                        rightWithIndex,
                                        "__HasMatch",
                                        each List.Contains(matchedRightIndexes, [__RightIndex])
                                    ),
                                    // For matched right rows, get corresponding left rows
                                    matchedPairs = Table.SelectRows(leftWithMatch, each [__MatchIndex] <> null),
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
                                        joined, "__LeftData", leftColumns, leftColumns
                                    ),
                                    cleaned = Table.RemoveColumns(expanded, {"__RightIndex", "__HasMatch"})
                                in
                                    cleaned
                            else
                                // Full join
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
                                        leftJoin, "__RightData", rightColumns, rightColumns
                                    ),
                                    // Find unmatched right rows
                                    matchedRightIndexes = List.Distinct(
                                        List.Select(Table.Column(leftWithMatch, "__MatchIndex"), each _ <> null)
                                    ),
                                    unmatchedRight = Table.SelectRows(
                                        rightWithIndex, each not List.Contains(matchedRightIndexes, [__RightIndex])
                                    ),
                                    // Create null columns for left side
                                    unmatchedRightWithNulls = List.Accumulate(
                                        leftColumns,
                                        unmatchedRight,
                                        (tbl, colName) => Table.AddColumn(tbl, colName, each null)
                                    ),
                                    // Combine both
                                    combined = Table.Combine({leftExpanded, unmatchedRightWithNulls}),
                                    cleaned = Table.RemoveColumns(
                                        combined, {"__LeftIndex", "__RightIndex", "__MatchIndex"}
                                    )
                                in
                                    cleaned
                    in
                        finalResult
        in
            result,
    IpMergeOnSubnetType = type function (
        leftTable as (
            type table meta [
                Documentation.FieldCaption = "Left Table",
                Documentation.FieldDescription = "Table containing IP addresses"
            ]
        ),
        leftIpColumn as (
            type text meta [
                Documentation.FieldCaption = "IP Column Name",
                Documentation.FieldDescription = "Name of the column containing IP addresses in the left table",
                Documentation.SampleValues = {"IP", "IPAddress", "Address"}
            ]
        ),
        rightTable as (
            type table meta [
                Documentation.FieldCaption = "Right Table",
                Documentation.FieldDescription = "Table containing subnet ranges"
            ]
        ),
        rightSubnetColumn as (
            type text meta [
                Documentation.FieldCaption = "Subnet Column Name",
                Documentation.FieldDescription = "Name of the column containing subnets in the right table. Supports both CIDR notation (e.g., 192.168.1.0/24) and dotted mask notation (e.g., 192.168.1.0 255.255.255.0).",
                Documentation.SampleValues = {"Subnet", "Network", "Range"}
            ]
        ),
        optional joinKind as (
            type text meta [
                Documentation.FieldCaption = "Join Kind",
                Documentation.FieldDescription = "Type of join to perform. Defaults to 'Left' if not specified.",
                Documentation.AllowedValues = {"Left", "Inner", "Right", "Full"}
            ]
        )
    ) as table meta [
        Documentation.Name = "IpMergeOnSubnet",
        Documentation.LongDescription = "Merges two tables based on subnet containment. Joins rows where an IP address in the left table falls within a subnet range in the right table. When multiple subnets contain an IP address, returns the most specific match (longest prefix). Similar to Table.NestedJoin but with subnet matching logic.",
        Documentation.Examples = {
            [
                Description = "Left join - keep all IPs, add subnet info where matched",
                Code = "IpCalc.IpMergeOnSubnet(ipTable, ""IP"", subnetTable, ""Subnet"", ""Left"")",
                Result = "Table with all rows from ipTable, subnet info added for matched IPs"
            ],
            [
                Description = "Inner join - only matched IPs",
                Code = "IpCalc.IpMergeOnSubnet(ipTable, ""IP"", subnetTable, ""Subnet"", ""Inner"")",
                Result = "Table with only IPs that fall within a subnet range"
            ]
        }
    ],
    IpMergeOnSubnetDoc = Value.ReplaceType(IpMergeOnSubnet, IpMergeOnSubnetType),
    //==============================================
    //   Table Functions
    //==============================================
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
        table as table, optional columnName as any, optional descending as logical, optional prefixOrder as text
    ) as table =>
        let
            // Determine column to sort by
            colName =
                if columnName is null then
                    Table.ColumnNames(table){0}
                else if Value.Is(columnName, type number) then
                    Table.ColumnNames(table){columnName}
                else
                    columnName,
            // Set defaults
            desc = if descending is null then false else descending,
            prefixOrd = if prefixOrder is null or prefixOrder = "" then "shortest" else prefixOrder,
            validPrefixOrders = {"none", "shortest", "longest"},
            prefixOrdValid =
                if List.Contains(validPrefixOrders, prefixOrd) then
                    prefixOrd
                else
                    error "Invalid prefixOrder. Must be one of: none, shortest, longest",
            // Parse IP/subnet and add temporary columns
            withParsed = Table.AddColumn(
                table, "_Parsed", each try _ParseIpNet(Record.Field(_, colName)) otherwise null
            ),
            withIpKey = Table.AddColumn(
                withParsed, "_IpKey", each if [_Parsed] = null then null else _IpToNbr([_Parsed][Ip])
            ),
            withPrefixKey = Table.AddColumn(
                withIpKey, "_PrefixKey", each if [_Parsed] = null then null else[_Parsed][Len]
            ),
            // Add column to handle nulls (empty cells and invalid entries go to end)
            withNullFlag = Table.AddColumn(
                withPrefixKey,
                "_IsNull",
                each let val = Record.Field(_, colName) in if val = null or val = "" or [_Parsed] = null then 1 else 0
            ),
            // Build sort columns dynamically
            sortColumns = {{"_IsNull", Order.Ascending}, {"_IpKey", if desc then Order.Descending else Order.Ascending}}
                & (
                    if prefixOrdValid <> "none" then
                        {{"_PrefixKey", if prefixOrdValid = "longest" then Order.Descending else Order.Ascending}}
                    else
                        {}
                ),
            // Sort table
            sorted = Table.Sort(withNullFlag, sortColumns),
            // Remove temporary columns
            result = Table.RemoveColumns(sorted, {"_Parsed", "_IpKey", "_PrefixKey", "_IsNull"})
        in
            result,
    IpSortTableType = type function (
        table as (
            type table meta [
                Documentation.FieldCaption = "Table",
                Documentation.FieldDescription = "Table to sort by IP address or subnet column"
            ]
        ),
        optional columnName as (
            type any meta [
                Documentation.FieldCaption = "Column Name",
                Documentation.FieldDescription = "Column name (text) or index (number) to sort by. Defaults to first column if not specified.",
                Documentation.SampleValues = {"IP", "Subnet", 0, 1}
            ]
        ),
        optional descending as (
            type logical meta [
                Documentation.FieldCaption = "Descending",
                Documentation.FieldDescription = "Sort in descending order. Defaults to false (ascending).",
                Documentation.SampleValues = {true, false}
            ]
        ),
        optional prefixOrder as (
            type text meta [
                Documentation.FieldCaption = "Prefix Order",
                Documentation.FieldDescription = "How to order entries with the same IP: 'shortest' (largest subnets first), 'longest' (smallest subnets first), or 'none' (ignore prefix). Defaults to 'shortest'.",
                Documentation.AllowedValues = {"none", "shortest", "longest"}
            ]
        )
    ) as table meta [
        Documentation.Name = "IpSortTable",
        Documentation.LongDescription = "Sorts a table by IP address or subnet column. Handles mixed IPv4 addresses and subnets. Empty values and invalid entries are placed at the end. Supports sorting by column name or index.",
        Documentation.Examples = {
            [
                Description = "Sort by column name, ascending, shortest prefix first",
                Code = "IpCalc.IpSortTable(myTable, ""IP"", false, ""shortest"")",
                Result = "Table sorted by IP column with largest subnets first"
            ],
            [
                Description = "Sort first column descending with default prefix order",
                Code = "IpCalc.IpSortTable(myTable, 0, true)",
                Result = "Table sorted by first column in descending order"
            ]
        }
    ],
    IpSortTableDoc = Value.ReplaceType(IpSortTable, IpSortTableType)
in
    //==============================================
    //   Exported Functions (with documentation)
    //==============================================
    [
        IpIsValid = IpIsValidDoc,
        IpStrToNbr = IpStrToNbrDoc,
        IpBinToStr = IpBinToStrDoc,
        IpAnd = IpAndDoc,
        IpOr = IpOrDoc,
        IpXor = IpXorDoc,
        IpWithoutMask = IpWithoutMaskDoc,
        IpSubnetLen = IpSubnetLenDoc,
        IpMask = IpMaskDoc,
        IpWildMask = IpWildMaskDoc,
        IpInvertMask = IpInvertMaskDoc,
        IpMaskLen = IpMaskLenDoc,
        IpAdd = IpAddDoc,
        IpDiff = IpDiffDoc,
        IpGetByte = IpGetByteDoc,
        IpSetByte = IpSetByteDoc,
        IpSubnetSize = IpSubnetSizeDoc,
        IpClearHostBits = IpClearHostBitsDoc,
        IpIsInSubnet = IpIsInSubnetDoc,
        IpSubnetIsInSubnet = IpSubnetIsInSubnetDoc,
        IpSubnetMatch = IpSubnetMatchDoc,
        IpSubnetVLookup = IpSubnetVLookupDoc,
        IpIsPrivate = IpIsPrivateDoc,
        IpDivideSubnet = IpDivideSubnetDoc,
        IpRangeToCIDR = IpRangeToCIDRDoc,
        IpSortArray = IpSortArrayDoc,
        IpSubnetSortArray = IpSubnetSortArrayDoc,
        IpFindOverlappingSubnets = IpFindOverlappingSubnetsDoc,
        IpSubnetAggregateArray = IpSubnetAggregateArrayDoc,
        IpSubtractSubnets = IpSubtractSubnetsDoc,
        IpCommonSubnets = IpCommonSubnetsDoc,
        IpMergeOnSubnet = IpMergeOnSubnetDoc,
        IpSortTable = IpSortTableDoc
    ]
