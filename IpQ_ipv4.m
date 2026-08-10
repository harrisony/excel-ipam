/*
    Converts a dotted-decimal IPv4 address to its base-256 integer form.

    The VBA reference accumulates one value for each dot-delimited segment:

        result = result * 256 + octet

    This M version makes the input contract explicit. It accepts exactly four
    ASCII decimal octets, each in the inclusive range 0..255. Leading zeroes
    are accepted because they do not change the numeric value.

    The JavaScript counterpart is ipStrToNbr, implemented through _ipToNbr;
    both use the same base-256 accumulation for valid IPv4 input. M rejects
    malformed or out-of-range input explicitly instead of inheriting
    JavaScript's numeric coercion behavior.

    Example:
        IpStrToBinDoc("1.2.3.4") = 16909060
*/
let
    IpStrToBin = (ip as nullable text) as number =>
        let
            fail = (message as text, detail as record) as none =>
                error Error.Record("IpStrToBin.InvalidInput", message, detail),

            parseOctet = (octet as text, position as number) as number =>
                let
                    asciiDigits = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"},
                    containsOnlyDigits = octet <> "" and Text.Select(octet, asciiDigits) = octet,
                    parsed =
                        if containsOnlyDigits then
                            Number.FromText(octet, "en-US")
                        else
                            null,
                    isValid =
                        if not containsOnlyDigits then
                            false
                        else
                            parsed >= 0
                                and parsed <= 255
                                and Number.RoundDown(parsed) = parsed
                in
                    if isValid then
                        parsed
                    else
                        fail(
                            "IPv4 octet must be an ASCII decimal integer from 0 through 255.",
                            [
                                Input = ip,
                                Component = "octet",
                                Position = position,
                                Value = octet,
                                Expected = "0..255"
                            ]
                        ),

            octets =
                if ip = null then
                    fail(
                        "IPv4 address cannot be null.",
                        [Input = ip, Component = "address", Expected = "text"]
                    )
                else
                    Text.Split(ip, "."),
            octetCount = List.Count(octets),
            parsedOctets =
                if octetCount <> 4 then
                    fail(
                        "IPv4 address must contain exactly four dot-delimited octets.",
                        [
                            Input = ip,
                            Component = "address",
                            Count = octetCount,
                            Expected = 4
                        ]
                    )
                else
                    List.Transform(
                        List.Positions(octets),
                        (position as number) => parseOctet(octets{position}, position + 1)
                    )
        in
            List.Accumulate(
                parsedOctets,
                0,
                (state as number, octet as number) => state * 256 + octet
            ),

    IpStrToBinType =
        type function (
            ip as (type nullable text meta [
                Documentation.Name = "IPv4 address",
                Documentation.Description = "A dotted-decimal IPv4 address with four octets."
            ])
        ) as number meta [
            Documentation.Name = "IpStrToBin",
            Documentation.Description = "Converts a dotted-decimal IPv4 address to its base-256 integer representation.",
            Documentation.LongDescription = "Returns the IPv4 address as a number in the range 0 through 4294967295. The input must contain exactly four ASCII decimal octets in the range 0 through 255.",
            Documentation.Examples = {
                [
                    Description = "Convert 1.2.3.4 to its integer representation.",
                    Code = "IpStrToBin(\"1.2.3.4\")",
                    Result = "16909060"
                ],
                [
                    Description = "Convert the first IPv4 address.",
                    Code = "IpStrToBin(\"0.0.0.0\")",
                    Result = "0"
                ],
                [
                    Description = "Convert the last IPv4 address.",
                    Code = "IpStrToBin(\"255.255.255.255\")",
                    Result = "4294967295"
                ]
            }
        ],

    IpStrToBinDoc = Value.ReplaceType(IpStrToBin, IpStrToBinType),

    /*
        Converts an IPv4 integer to dotted-decimal text.

        The VBA reference extracts four base-256 digits from right to left,
        then prepends each digit to the result. The validated M contract is
        the unsigned IPv4 integer range 0..4294967295.

        The JavaScript counterpart is ipNbrToStr, implemented through _ipFromNbr;
        this function preserves its four-octet output for the validated IPv4
        range while making the range and whole-number requirements explicit.

        Example:
            IpBinToStrDoc(16909060) = "1.2.3.4"
    */
    IpBinToStr = (ip as nullable number) as text =>
        let
            fail = (message as text, detail as record) as none =>
                error Error.Record("IpBinToStr.InvalidInput", message, detail),

            checkedIp =
                if ip = null then
                    fail(
                        "IPv4 integer cannot be null.",
                        [Input = ip, Component = "address", Expected = "number"]
                    )
                else if ip < 0 or ip > 4294967295 then
                    fail(
                        "IPv4 integer must be in the range 0 through 4294967295.",
                        [
                            Input = ip,
                            Component = "address",
                            Expected = "0..4294967295"
                        ]
                    )
                else if Number.RoundDown(ip) <> ip then
                    fail(
                        "IPv4 integer must be a whole number.",
                        [
                            Input = ip,
                            Component = "address",
                            Expected = "whole number"
                        ]
                    )
                else
                    ip,

            leastSignificantFirst =
                List.Accumulate(
                    {1..4},
                    [Value = checkedIp, Octets = {}],
                    (state as record, _ as number) as record =>
                        let
                            quotient = Number.IntegerDivide(state[Value], 256),
                            remainder = Number.Mod(state[Value], 256)
                        in
                            [
                                Value = quotient,
                                Octets = state[Octets] & {remainder}
                            ]
                )[Octets],
            octets = List.Reverse(leastSignificantFirst)
        in
            Text.Combine(
                List.Transform(
                    octets,
                    (octet as number) => Number.ToText(octet, "0", "en-US")
                ),
                "."
            ),

    IpBinToStrType =
        type function (
            ip as (type nullable number meta [
                Documentation.Name = "IPv4 integer",
                Documentation.Description = "An integer in the range 0 through 4294967295."
            ])
        ) as text meta [
            Documentation.Name = "IpBinToStr",
            Documentation.Description = "Converts an IPv4 integer to dotted-decimal text.",
            Documentation.LongDescription = "Returns a four-octet IPv4 address for an integer in the range 0 through 4294967295.",
            Documentation.Examples = {
                [
                    Description = "Convert 16909060 to dotted-decimal IPv4 text.",
                    Code = "IpBinToStr(16909060)",
                    Result = "1.2.3.4"
                ],
                [
                    Description = "Convert the first IPv4 address.",
                    Code = "IpBinToStr(0)",
                    Result = "0.0.0.0"
                ],
                [
                    Description = "Convert the last IPv4 address.",
                    Code = "IpBinToStr(4294967295)",
                    Result = "255.255.255.255"
                ]
            }
        ],

    IpBinToStrDoc = Value.ReplaceType(IpBinToStr, IpBinToStrType),

    /*
        Parses the rightmost byte from an IPv4 address fragment.

        The VBA reference removes the rightmost dot-delimited component from
        its ByRef input and returns that component as an integer. M values are
        immutable, so this version returns both values in a record. The
        remainder is an empty text value when the input has no dot.

        Unlike VBA's permissive Val conversion, the M contract requires the
        extracted component to be a non-empty ASCII decimal integer from 0
        through 255. Earlier components remain in Remainder for a subsequent
        parse, matching the reference's right-to-left iteration model.

        Examples:
            IpParseDoc("192.168.1.32")
                = [Byte = 32, Remainder = "192.168.1"]
            IpParseDoc("32")
                = [Byte = 32, Remainder = ""]
    */
    IpParse = (ip as nullable text) as record =>
        let
            fail = (message as text, detail as record) as none =>
                error Error.Record("IpParse.InvalidInput", message, detail),

            checkedIp =
                if ip = null then
                    fail(
                        "IPv4 address fragment cannot be null.",
                        [Input = ip, Component = "address fragment", Expected = "text"]
                    )
                else
                    ip,
            separatorPosition = Text.PositionOf(checkedIp, ".", Occurrence.Last),
            byteText =
                if separatorPosition < 0 then
                    checkedIp
                else
                    Text.Range(checkedIp, separatorPosition + 1),
            remainder =
                if separatorPosition < 0 then
                    ""
                else
                    Text.Start(checkedIp, separatorPosition),
            asciiDigits = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"},
            containsOnlyDigits = byteText <> "" and Text.Select(byteText, asciiDigits) = byteText,
            byteValue =
                if containsOnlyDigits then
                    Number.FromText(byteText, "en-US")
                else
                    null,
            isValidByte =
                if not containsOnlyDigits then
                    false
                else
                    byteValue >= 0
                        and byteValue <= 255
                        and Number.RoundDown(byteValue) = byteValue
        in
            if isValidByte then
                [Byte = byteValue, Remainder = remainder]
            else
                fail(
                    "IPv4 byte must be an ASCII decimal integer from 0 through 255.",
                    [
                        Input = ip,
                        Component = "byte",
                        Value = byteText,
                        Expected = "0..255"
                    ]
                ),

    IpParseType =
        type function (
            ip as (type nullable text meta [
                Documentation.Name = "IPv4 address fragment",
                Documentation.Description = "An IPv4 address or partial address whose rightmost byte is parsed."
            ])
        ) as [Byte = number, Remainder = text] meta [
            Documentation.Name = "IpParse",
            Documentation.Description = "Parses and removes the rightmost byte from an IPv4 address fragment.",
            Documentation.LongDescription = "Returns a record containing Byte and Remainder fields. The input is processed from right to left; Remainder is the text before the final dot, or an empty text value when no dot is present. The extracted byte must be an ASCII decimal integer from 0 through 255.",
            Documentation.Examples = {
                [
                    Description = "Parse the rightmost byte and retain the preceding address fragment.",
                    Code = "IpParse(\"192.168.1.32\")",
                    Result = "[Byte = 32, Remainder = \"192.168.1\"]"
                ],
                [
                    Description = "Parse a single byte with no preceding address fragment.",
                    Code = "IpParse(\"32\")",
                    Result = "[Byte = 32, Remainder = \"\"]"
                ],
                [
                    Description = "Parse the largest valid IPv4 byte.",
                    Code = "IpParse(\"255\")",
                    Result = "[Byte = 255, Remainder = \"\"]"
                ]
            }
        ],

    IpParseDoc = Value.ReplaceType(IpParse, IpParseType),

    /*
        Prepends the low byte of an IPv4 calculation to an address fragment.

        VBA returns the remaining base-256 carry while mutating its ByRef
        address argument. M values are immutable, so this version returns a
        record containing both results. The low eight bits are represented by
        Ip and the remaining quotient by Carry. A non-empty address fragment
        is separated with a dot; an empty fragment receives no leading dot.

        Examples:
            IpBuildDoc(192, "168.1.1")
                = [Ip = "192.168.1.1", Carry = 0]
            IpBuildDoc(258, "1")
                = [Ip = "2.1", Carry = 1]

        Unlike VBA's implicit numeric coercion, the M contract requires a
        non-negative whole number that can be represented exactly by M's
        double-precision number type.
    */
    IpBuild = (ipByte as nullable number, ip as nullable text) as record =>
        let
            fail = (message as text, detail as record) as none =>
                error Error.Record("IpBuild.InvalidInput", message, detail),

            checkedIpByte =
                if ipByte = null then
                    fail(
                        "IPv4 byte value cannot be null.",
                        [Input = ipByte, Component = "byte", Expected = "non-negative whole number"]
                    )
                else if ipByte < 0 then
                    fail(
                        "IPv4 byte value must be non-negative.",
                        [Input = ipByte, Component = "byte", Expected = "non-negative whole number"]
                    )
                else if Number.RoundDown(ipByte) <> ipByte then
                    fail(
                        "IPv4 byte value must be a whole number.",
                        [Input = ipByte, Component = "byte", Expected = "non-negative whole number"]
                    )
                else if ipByte > Number.Power(2, 53) - 1 then
                    fail(
                        "IPv4 byte value must be within M's exact integer range.",
                        [Input = ipByte, Component = "byte", Expected = "0..9007199254740991"]
                    )
                else
                    ipByte,

            checkedIp =
                if ip = null then
                    fail(
                        "IPv4 address fragment cannot be null.",
                        [Input = ip, Component = "address fragment", Expected = "text"]
                    )
                else
                    ip,
            carry = Number.IntegerDivide(checkedIpByte, 256),
            lowByte = Number.Mod(checkedIpByte, 256),
            prefix = Number.ToText(lowByte, "0", "en-US"),
            builtIp = if checkedIp = "" then prefix else prefix & "." & checkedIp
        in
            [Ip = builtIp, Carry = carry],

    IpBuildType =
        type function (
            ipByte as (type nullable number meta [
                Documentation.Name = "IPv4 byte value",
                Documentation.Description = "A non-negative whole number whose low eight bits are prepended."
            ]),
            ip as (type nullable text meta [
                Documentation.Name = "IPv4 address fragment",
                Documentation.Description = "The existing dotted-decimal fragment to receive the new low byte."
            ])
        ) as [Ip = text, Carry = number] meta [
            Documentation.Name = "IpBuild",
            Documentation.Description = "Prepends the low byte of a number to an IPv4 address fragment and returns the remaining carry.",
            Documentation.LongDescription = "Returns a record with Ip and Carry fields. Ip contains the low eight bits followed by the original fragment, and Carry is the integer quotient after division by 256. This record is the immutable M equivalent of the VBA function's ByRef string and numeric return value.",
            Documentation.Examples = {
                [
                    Description = "Prepend a byte without a carry.",
                    Code = "IpBuild(192, \"168.1.1\")",
                    Result = "[Ip = \"192.168.1.1\", Carry = 0]"
                ],
                [
                    Description = "Prepend the low byte and return the carry.",
                    Code = "IpBuild(258, \"1\")",
                    Result = "[Ip = \"2.1\", Carry = 1]"
                ],
                [
                    Description = "Build the first byte of an address.",
                    Code = "IpBuild(255, \"\")",
                    Result = "[Ip = \"255\", Carry = 0]"
                ]
            }
        ],

    IpBuildDoc = Value.ReplaceType(IpBuild, IpBuildType),

    /*
        Compares the first n bits of two IPv4 addresses.

        The VBA reference compares complete octets from left to right and,
        when n ends inside an octet, compares only that octet's leading bits.
        This M implementation uses the validated IPv4 integer representation
        and integer-dividing away the host bits. A zero-bit comparison therefore
        returns true for any two valid IPv4 addresses.

        The JavaScript reference has no direct public IpComp function. Its
        IpNet.matchIp and IpNet.matchSubnet operations provide the related
        subnet-prefix behavior, while this function preserves the VBA API's
        arbitrary prefix length.

        Examples:
            IpCompDoc("10.0.0.0", "10.1.0.0", 9) = true
            IpCompDoc("10.0.0.0", "10.1.0.0", 16) = false
            IpCompDoc("192.168.1.1", "10.0.0.1", 0) = true

        Unlike VBA's implicit Val conversion, both addresses must be valid
        dotted-decimal IPv4 text and n must be a whole number from 0 through
        32. Invalid inputs raise the structured validation errors from
        IpStrToBin or the explicit prefix validation below.
    */
    IpComp = (ip1 as nullable text, ip2 as nullable text, n as nullable number) as logical =>
        let
            fail = (message as text, detail as record) as none =>
                error Error.Record("IpComp.InvalidInput", message, detail),

            prefixLength =
                if n = null then
                    fail(
                        "Prefix length cannot be null.",
                        [Input = n, Component = "prefix length", Expected = "whole number 0..32"]
                    )
                else if n < 0 or n > 32 or Number.RoundDown(n) <> n then
                    fail(
                        "Prefix length must be a whole number from 0 through 32.",
                        [Input = n, Component = "prefix length", Expected = "whole number 0..32"]
                    )
                else
                    n,

            firstAddress = IpStrToBin(ip1),
            secondAddress = IpStrToBin(ip2),
            hostBitCount = 32 - prefixLength,
            divisor = Number.Power(2, hostBitCount),
            firstPrefix = Number.IntegerDivide(firstAddress, divisor),
            secondPrefix = Number.IntegerDivide(secondAddress, divisor)
        in
            firstPrefix = secondPrefix,

    IpCompType =
        type function (
            ip1 as (type nullable text meta [
                Documentation.Name = "First IPv4 address",
                Documentation.Description = "A dotted-decimal IPv4 address with four octets."
            ]),
            ip2 as (type nullable text meta [
                Documentation.Name = "Second IPv4 address",
                Documentation.Description = "A dotted-decimal IPv4 address with four octets."
            ]),
            n as (type nullable number meta [
                Documentation.Name = "Prefix length",
                Documentation.Description = "The number of leading bits to compare, from 0 through 32."
            ])
        ) as logical meta [
            Documentation.Name = "IpComp",
            Documentation.Description = "Compares the first n bits of two IPv4 addresses.",
            Documentation.LongDescription = "Returns true when the first n bits of both validated dotted-decimal IPv4 addresses are equal. Prefix lengths from 0 through 32 are supported; a zero-bit comparison returns true. Invalid addresses or prefix lengths raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Compare a prefix that ends inside the second octet.",
                    Code = "IpComp(\"10.0.0.0\", \"10.1.0.0\", 9)",
                    Result = "true"
                ],
                [
                    Description = "Show that the same addresses differ in their first 16 bits.",
                    Code = "IpComp(\"10.0.0.0\", \"10.1.0.0\", 16)",
                    Result = "false"
                ],
                [
                    Description = "Compare zero leading bits.",
                    Code = "IpComp(\"192.168.1.1\", \"10.0.0.1\", 0)",
                    Result = "true"
                ]
            }
        ],

    IpCompDoc = Value.ReplaceType(IpComp, IpCompType),

    /*
        Returns the prefix length represented by a dotted-decimal IPv4 mask.

        A subnet mask is valid when it is the canonical mask for one of the
        prefix lengths 0..32. Comparing against those 33 masks is explicit,
        exact for IPv4 numbers, and rejects non-contiguous masks rather than
        assigning them an ambiguous prefix length.

        The JavaScript counterpart is ipMaskLen. Unlike its binary-string
        implementation, this M contract rejects non-contiguous masks explicitly
        instead of deriving a misleading prefix length from the last one bit.

        Example:
            IpMaskLenDoc("255.255.255.0") = 24
    */
    IpMaskLen = (mask as nullable text) as number =>
        let
            fail = (message as text, detail as record) as none =>
                error Error.Record("IpMaskLen.InvalidInput", message, detail),

            maskValue =
                if mask = null then
                    fail(
                        "IPv4 mask cannot be null.",
                        [Input = mask, Component = "mask", Expected = "dotted-decimal IPv4 mask"]
                    )
                else
                    IpStrToBin(mask),
            prefixLengths = {0..32},
            canonicalMasks =
                List.Transform(
                    prefixLengths,
                    (prefix as number) =>
                        (Number.Power(2, prefix) - 1) * Number.Power(2, 32 - prefix)
                ),
            prefixLength = List.PositionOf(canonicalMasks, maskValue)
        in
            if prefixLength >= 0 then
                prefixLength
            else
                fail(
                    "IPv4 mask must contain contiguous one bits followed by contiguous zero bits.",
                    [
                        Input = mask,
                        Component = "mask",
                        Value = maskValue,
                        Expected = "canonical IPv4 netmask"
                    ]
                ),

    IpMaskLenType =
        type function (
            mask as (type nullable text meta [
                Documentation.Name = "IPv4 mask",
                Documentation.Description = "A dotted-decimal contiguous IPv4 subnet mask."
            ])
        ) as number meta [
            Documentation.Name = "IpMaskLen",
            Documentation.Description = "Returns the prefix length represented by a dotted-decimal IPv4 mask.",
            Documentation.LongDescription = "Accepts canonical contiguous IPv4 subnet masks and returns a prefix length from 0 through 32.",
            Documentation.Examples = {
                [
                    Description = "Get the prefix length of a /24 mask.",
                    Code = "IpMaskLen(\"255.255.255.0\")",
                    Result = "24"
                ],
                [
                    Description = "Get the prefix length of the default route mask.",
                    Code = "IpMaskLen(\"0.0.0.0\")",
                    Result = "0"
                ],
                [
                    Description = "Get the prefix length of the host mask.",
                    Code = "IpMaskLen(\"255.255.255.255\")",
                    Result = "32"
                ]
            }
        ],

    IpMaskLenDoc = Value.ReplaceType(IpMaskLen, IpMaskLenType),

    /*
        Parses an IPv4 address with an optional subnet mask.

        VBA modifies its ByRef string argument to remove the mask. M values
        are immutable, so this version returns a record containing both the
        mask-free address and its prefix length. It accepts CIDR notation
        (address/prefix) and dotted-mask notation (address mask); an address
        without a mask is treated as a /32 host address.

        Examples:
            IpSubnetParseDoc("192.168.1.1/24")
                = [Address = "192.168.1.1", PrefixLength = 24]
            IpSubnetParseDoc("192.168.1.1 255.255.255.0")
                = [Address = "192.168.1.1", PrefixLength = 24]
    */
    IpSubnetParse = (subnet as nullable text) as record =>
        let
            fail = (message as text, detail as record) as none =>
                error Error.Record("IpSubnetParse.InvalidInput", message, detail),

            parsePrefix = (prefixText as text) as number =>
                let
                    asciiDigits = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"},
                    containsOnlyDigits = prefixText <> "" and Text.Select(prefixText, asciiDigits) = prefixText,
                    prefix =
                        if containsOnlyDigits then
                            Number.FromText(prefixText, "en-US")
                        else
                            null,
                    isValid =
                        if not containsOnlyDigits then
                            false
                        else
                            prefix >= 0
                                and prefix <= 32
                                and Number.RoundDown(prefix) = prefix
                in
                    if isValid then
                        prefix
                    else
                        fail(
                            "IPv4 prefix length must be an integer from 0 through 32.",
                            [
                                Input = subnet,
                                Component = "prefix length",
                                Value = prefixText,
                                Expected = "0..32"
                            ]
                        ),

            checkedSubnet =
                if subnet = null then
                    fail(
                        "IPv4 subnet cannot be null.",
                        [Input = subnet, Component = "subnet", Expected = "IPv4 subnet text"]
                    )
                else
                    subnet,
            slashPosition = Text.PositionOf(checkedSubnet, "/"),
            spacePosition = Text.PositionOf(checkedSubnet, " "),
            hasSlash = slashPosition >= 0,
            hasSpace = spacePosition >= 0,
            addressText =
                if hasSlash then
                    Text.Start(checkedSubnet, slashPosition)
                else if hasSpace then
                    Text.Start(checkedSubnet, spacePosition)
                else
                    checkedSubnet,
            suffixText =
                if hasSlash then
                    Text.Range(checkedSubnet, slashPosition + 1)
                else if hasSpace then
                    Text.Range(checkedSubnet, spacePosition + 1)
                else
                    null,
            prefixLength =
                if hasSlash then
                    parsePrefix(suffixText)
                else if hasSpace then
                    IpMaskLen(suffixText)
                else
                    32,
            validatedAddress = try IpStrToBin(addressText),
            result =
                if validatedAddress[HasError] then
                    fail(
                        "IPv4 address must be in the dotted-decimal format.",
                        [
                            Input = subnet,
                            Component = "address",
                            Value = addressText,
                            Cause = validatedAddress[Error]
                        ]
                    )
                else
                    [Address = addressText, PrefixLength = prefixLength]
        in
            result,

    IpSubnetParseType =
        type function (
            subnet as (type nullable text meta [
                Documentation.Name = "IPv4 subnet",
                Documentation.Description = "An IPv4 address with optional CIDR or dotted-mask notation."
            ])
        ) as [Address = text, PrefixLength = number] meta [
            Documentation.Name = "IpSubnetParse",
            Documentation.Description = "Parses an IPv4 address and optional subnet mask.",
            Documentation.LongDescription = "Returns a record with Address and PrefixLength fields. CIDR notation and dotted-decimal mask notation are supported; an address without a mask is treated as /32.",
            Documentation.Examples = {
                [
                    Description = "Parse CIDR notation.",
                    Code = "IpSubnetParse(\"192.168.1.1/24\")",
                    Result = "[Address = \"192.168.1.1\", PrefixLength = 24]"
                ],
                [
                    Description = "Parse dotted-mask notation.",
                    Code = "IpSubnetParse(\"192.168.1.1 255.255.255.0\")",
                    Result = "[Address = \"192.168.1.1\", PrefixLength = 24]"
                ],
                [
                    Description = "Treat an address without a mask as a /32 host address.",
                    Code = "IpSubnetParse(\"192.168.1.1\")",
                    Result = "[Address = \"192.168.1.1\", PrefixLength = 32]"
                ]
            }
        ],

    IpSubnetParseDoc = Value.ReplaceType(IpSubnetParse, IpSubnetParseType),

    /*
        Returns the prefix length from an IPv4 subnet expression.

        The VBA implementation finds a CIDR prefix after "/", derives a
        prefix from a dotted-decimal mask after a space, and treats an
        unmasked address as a /32 host address. M values are immutable, so
        this function delegates to IpSubnetParse and selects its
        PrefixLength field rather than re-parsing the subnet independently.

        The JavaScript counterpart is ipSubnetLen, which reads the len field
        from IpNet. This M contract preserves the documented results while
        explicitly rejecting null, malformed addresses, invalid prefix
        lengths, and non-canonical masks through IpSubnetParse.

        Examples:
            IpSubnetLenDoc("192.168.1.1/24") = 24
            IpSubnetLenDoc("192.168.1.1 255.255.255.0") = 24
            IpSubnetLenDoc("192.168.1.1") = 32
    */
    IpSubnetLen = (subnet as nullable text) as number =>
        let
            parsedSubnet = IpSubnetParse(subnet)
        in
            parsedSubnet[PrefixLength],

    IpSubnetLenType =
        type function (
            subnet as (type nullable text meta [
                Documentation.Name = "IPv4 subnet",
                Documentation.Description = "An IPv4 address with optional CIDR or dotted-mask notation."
            ])
        ) as number meta [
            Documentation.Name = "IpSubnetLen",
            Documentation.Description = "Returns the prefix length from an IPv4 subnet expression.",
            Documentation.LongDescription = "Accepts a dotted IPv4 address, CIDR notation, or dotted-decimal mask notation and returns a prefix length from 0 through 32. An unmasked address is treated as /32; malformed addresses, invalid prefixes, and non-canonical masks raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Get the prefix length from CIDR notation.",
                    Code = "IpSubnetLen(\"192.168.1.1/24\")",
                    Result = "24"
                ],
                [
                    Description = "Get the prefix length from dotted-mask notation.",
                    Code = "IpSubnetLen(\"192.168.1.1 255.255.255.0\")",
                    Result = "24"
                ],
                [
                    Description = "Treat an address without a mask as a /32 host address.",
                    Code = "IpSubnetLen(\"192.168.1.1\")",
                    Result = "32"
                ]
            }
        ],

    IpSubnetLenDoc = Value.ReplaceType(IpSubnetLen, IpSubnetLenType)
in
    [
        IpStrToBin = IpStrToBinDoc,
        IpBinToStr = IpBinToStrDoc,
        IpParse = IpParseDoc,
        IpBuild = IpBuildDoc,
        IpComp = IpCompDoc,
        IpMaskLen = IpMaskLenDoc,
        IpSubnetParse = IpSubnetParseDoc,
        IpSubnetLen = IpSubnetLenDoc
    ]
