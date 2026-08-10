/*
    Converts a dotted-decimal IPv4 address to its base-256 integer form.

    The VBA reference accumulates one value for each dot-delimited segment:

        result = result * 256 + octet

    This M version makes the input contract explicit. It accepts exactly four
    ASCII decimal octets, each in the inclusive range 0..255. Leading zeroes
    are accepted because they do not change the numeric value.

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
                            quotient = Number.RoundDown(state[Value] / 256),
                            remainder = state[Value] - quotient * 256
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
        ) as record meta [
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
        Returns the prefix length represented by a dotted-decimal IPv4 mask.

        A subnet mask is valid when it is the canonical mask for one of the
        prefix lengths 0..32. Comparing against those 33 masks is explicit,
        exact for IPv4 numbers, and rejects non-contiguous masks rather than
        assigning them an ambiguous prefix length.

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
            addressValue = IpStrToBin(addressText),
            result =
                if addressValue >= 0 then
                    [Address = addressText, PrefixLength = prefixLength]
                else
                    fail(
                        "IPv4 address must be in the dotted-decimal format.",
                        [Input = subnet, Component = "address", Value = addressText]
                    )
        in
            result,

    IpSubnetParseType =
        type function (
            subnet as (type nullable text meta [
                Documentation.Name = "IPv4 subnet",
                Documentation.Description = "An IPv4 address with optional CIDR or dotted-mask notation."
            ])
        ) as record meta [
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

    IpSubnetParseDoc = Value.ReplaceType(IpSubnetParse, IpSubnetParseType)
in
    [
        IpStrToBin = IpStrToBinDoc,
        IpBinToStr = IpBinToStrDoc,
        IpParse = IpParseDoc,
        IpMaskLen = IpMaskLenDoc,
        IpSubnetParse = IpSubnetParseDoc
    ]
